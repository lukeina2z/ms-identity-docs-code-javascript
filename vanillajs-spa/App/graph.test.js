/**
 * @jest-environment jsdom
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, 'public', 'index.html'), 'utf8');
const scripts = ['authConfig.js', 'ui.js', 'claimUtils.js', 'auth.js'].map((name) =>
    fs.readFileSync(path.join(__dirname, 'public', name), 'utf8')
);

class InteractionRequiredAuthError extends Error {}

const account = {
    homeAccountId: 'test-account',
    username: 'user@example.com',
    idTokenClaims: { name: 'Test User', aud: 'test-client' }
};

const profile = {
    '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users/$entity',
    id: 'graph-user',
    displayName: 'Test User',
    mail: null,
    userPrincipalName: 'user@example.com'
};

let client;
let context;

async function initializeApp() {
    for (const script of scripts) {
        await vm.runInContext(script, context);
    }
}

beforeEach(() => {
    document.documentElement.innerHTML = html;
    client = {
        initialize: jest.fn().mockResolvedValue(),
        handleRedirectPromise: jest.fn().mockResolvedValue(null),
        getActiveAccount: jest.fn().mockReturnValue(null),
        getAllAccounts: jest.fn().mockReturnValue([]),
        setActiveAccount: jest.fn(),
        loginPopup: jest.fn().mockResolvedValue({ account, idToken: 'not-an-access-token' }),
        loginRedirect: jest.fn().mockResolvedValue(),
        acquireTokenSilent: jest.fn().mockResolvedValue({ accessToken: 'graph-access-token' }),
        acquireTokenRedirect: jest.fn().mockResolvedValue()
    };
    context = vm.createContext({
        document,
        window,
        msal: {
            PublicClientApplication: jest.fn().mockImplementation(() => client),
            InteractionRequiredAuthError
        },
        console: { error: jest.fn(), warn: jest.fn() },
        fetch: jest.fn().mockResolvedValue({
            ok: true,
            json: jest.fn().mockResolvedValue(profile)
        })
    });
});

test('keeps the profile hidden before sign-in and places it above the claims table', async () => {
    await initializeApp();

    const profileDiv = document.getElementById('profile-div');
    expect(profileDiv.classList.contains('d-none')).toBe(true);
    expect(profileDiv.compareDocumentPosition(document.getElementById('table-div')) &
        Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(context.fetch).not.toHaveBeenCalled();
});

test('popup sign-in acquires a Graph access token for the selected account and renders /me', async () => {
    await initializeApp();
    await context.signIn('popup');

    expect(client.loginPopup).toHaveBeenCalledWith({
        scopes: ['User.Read'],
        redirectUri: '/redirect'
    });
    expect(client.acquireTokenSilent).toHaveBeenCalledWith({
        scopes: ['User.Read'],
        account
    });
    expect(context.fetch).toHaveBeenCalledTimes(1);
    expect(context.fetch).toHaveBeenCalledWith('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: 'Bearer graph-access-token' }
    });
    expect(document.getElementById('profile-result').textContent).toBe(JSON.stringify(profile, null, 2));
    expect(document.getElementById('profile-result').classList.contains('d-none')).toBe(false);
    expect(document.getElementById('profile-status').classList.contains('d-none')).toBe(true);
    expect(document.getElementById('table-body-div').rows).toHaveLength(2);
});

test('redirect sign-in requests User.Read', async () => {
    await initializeApp();
    await context.signIn('redirect');

    expect(client.loginRedirect).toHaveBeenCalledWith({ scopes: ['User.Read'] });
});

test('loads the profile after a redirect response', async () => {
    client.handleRedirectPromise.mockResolvedValue({ account });
    await initializeApp();

    expect(client.setActiveAccount).toHaveBeenCalledWith(account);
    expect(context.fetch).toHaveBeenCalledTimes(1);
    expect(document.getElementById('profile-div').classList.contains('d-none')).toBe(false);
});

test('restores a single cached account and renders claims only once', async () => {
    client.getAllAccounts.mockReturnValue([account]);
    await initializeApp();

    expect(client.setActiveAccount).toHaveBeenCalledWith(account);
    expect(document.getElementById('welcome-div').textContent).toBe('Welcome user@example.com!');
    expect(document.getElementById('table-body-div').rows).toHaveLength(2);
    expect(context.fetch).toHaveBeenCalledTimes(1);
});

test('uses the active cached account even when multiple accounts exist', async () => {
    client.getActiveAccount.mockReturnValue(account);
    client.getAllAccounts.mockReturnValue([account, { homeAccountId: 'other' }]);
    await initializeApp();

    expect(client.acquireTokenSilent).toHaveBeenCalledWith({ scopes: ['User.Read'], account });
});

test('does not guess an account when multiple accounts have no active selection', async () => {
    client.getAllAccounts.mockReturnValue([account, { homeAccountId: 'other' }]);
    await initializeApp();

    expect(context.console.warn).toHaveBeenCalledWith('Multiple accounts detected.');
    expect(context.fetch).not.toHaveBeenCalled();
});

test('shows a loading state while the profile request is pending', async () => {
    let resolveResponse;
    const responsePromise = new Promise((resolve) => {
        resolveResponse = resolve;
    });
    context.fetch.mockReturnValue(responsePromise);
    await initializeApp();
    const pending = context.handleResponse({ account });

    expect(document.getElementById('profile-status').textContent).toContain('Loading');
    expect(document.getElementById('profile-result').classList.contains('d-none')).toBe(true);

    resolveResponse({ ok: true, json: async () => profile });
    await pending;
    expect(document.getElementById('profile-status').classList.contains('d-none')).toBe(true);
});

test('falls back to redirect only when interaction is required', async () => {
    client.acquireTokenSilent.mockRejectedValue(new InteractionRequiredAuthError('Consent required'));
    await initializeApp();
    await context.signIn('popup');

    expect(client.acquireTokenRedirect).toHaveBeenCalledWith({ scopes: ['User.Read'], account });
    expect(context.fetch).not.toHaveBeenCalled();
    expect(document.getElementById('profile-status').textContent).toContain('Loading');
});

test('shows redirect acquisition failures without removing ID-token claims', async () => {
    client.acquireTokenSilent.mockRejectedValue(new InteractionRequiredAuthError('Consent required'));
    client.acquireTokenRedirect.mockRejectedValue(new Error('Redirect unavailable'));
    await initializeApp();
    await context.signIn('popup');

    expect(context.fetch).not.toHaveBeenCalled();
    expect(document.getElementById('profile-status').textContent).toContain('Redirect unavailable');
    expect(document.getElementById('table-body-div').rows).toHaveLength(2);
});

test('surfaces non-interaction token errors without starting a redirect', async () => {
    client.acquireTokenSilent.mockRejectedValue(new Error('Token acquisition failed'));
    await initializeApp();
    await context.signIn('popup');

    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
    expect(context.fetch).not.toHaveBeenCalled();
    expect(document.getElementById('profile-status').textContent).toContain('Token acquisition failed');
    expect(context.console.error).toHaveBeenCalled();
});

test('does not call Graph with a missing access token', async () => {
    client.acquireTokenSilent.mockResolvedValue({ accessToken: '' });
    await initializeApp();
    await context.signIn('popup');

    expect(context.fetch).not.toHaveBeenCalled();
    expect(document.getElementById('profile-status').textContent).toContain('did not return');
});

test('shows the HTTP status and Graph error details while preserving claims', async () => {
    context.fetch.mockResolvedValue({
        ok: false,
        status: 403,
        text: async () => '{"error":{"message":"Insufficient privileges"}}'
    });
    await initializeApp();
    await context.signIn('popup');

    const status = document.getElementById('profile-status');
    expect(status.textContent).toContain('403');
    expect(status.textContent).toContain('Insufficient privileges');
    expect(status.classList.contains('text-danger')).toBe(true);
    expect(document.getElementById('table-div').classList.contains('d-none')).toBe(false);
});

test('shows network failures instead of an empty successful profile', async () => {
    context.fetch.mockRejectedValue(new Error('Failed to fetch'));
    await initializeApp();
    await context.signIn('popup');

    expect(document.getElementById('profile-status').textContent).toContain('Failed to fetch');
    expect(document.getElementById('profile-result').classList.contains('d-none')).toBe(true);
});

test('shows invalid JSON response errors', async () => {
    context.fetch.mockResolvedValue({
        ok: true,
        json: async () => { throw new Error('Invalid JSON response'); }
    });
    await initializeApp();
    await context.signIn('popup');

    expect(document.getElementById('profile-status').textContent).toContain('Invalid JSON response');
});

test('renders returned profile values as text rather than HTML', async () => {
    const untrustedProfile = { ...profile, displayName: '<img src=x onerror=alert(1)>' };
    context.fetch.mockResolvedValue({ ok: true, json: async () => untrustedProfile });
    await initializeApp();
    await context.signIn('popup');

    const result = document.getElementById('profile-result');
    expect(result.textContent).toContain(untrustedProfile.displayName);
    expect(result.querySelector('img')).toBeNull();
});

test('clears stale profile data on a failed reload and does not duplicate claims', async () => {
    await initializeApp();
    await context.signIn('popup');
    context.fetch.mockRejectedValue(new Error('Profile refresh failed'));
    await context.handleResponse({ account });

    expect(document.getElementById('profile-result').textContent).toBe('');
    expect(document.getElementById('profile-status').textContent).toContain('Profile refresh failed');
    expect(document.getElementById('table-body-div').rows).toHaveLength(2);
});
