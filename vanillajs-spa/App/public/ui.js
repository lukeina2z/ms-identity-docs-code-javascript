// Select DOM elements to work with
const signInButton = document.getElementById('signIn');
const signOutButton = document.getElementById('signOut');
const titleDiv = document.getElementById('title-div');
const welcomeDiv = document.getElementById('welcome-div');
const tableDiv = document.getElementById('table-div');
const tableBody = document.getElementById('table-body-div');
const profileDiv = document.getElementById('profile-div');
const profileStatus = document.getElementById('profile-status');
const profileResult = document.getElementById('profile-result');

function showWelcomeMessage(account) {
    signInButton.classList.add('d-none');
    signOutButton.classList.remove('d-none');
    titleDiv.classList.add('d-none');
    welcomeDiv.classList.remove('d-none');
    welcomeDiv.textContent = `Welcome ${account.username}!`;
    updateTable(account);
};

function showProfileLoading() {
    profileDiv.classList.remove('d-none');
    profileStatus.classList.remove('d-none', 'text-danger');
    profileStatus.textContent = 'Loading Microsoft Graph user profile...';
    profileResult.classList.add('d-none');
    profileResult.textContent = '';
}

function showUserProfile(profile) {
    profileStatus.classList.add('d-none');
    profileStatus.textContent = '';
    profileResult.textContent = JSON.stringify(profile, null, 2);
    profileResult.classList.remove('d-none');
}

function showProfileError(error) {
    profileStatus.classList.remove('d-none');
    profileStatus.classList.add('text-danger');
    profileStatus.textContent = `Unable to load Microsoft Graph user profile: ${String(error)}`;
}

function updateTable(account) {
    tableDiv.classList.remove('d-none');
    tableBody.replaceChildren();
    
    const tokenClaims = createClaimsTable(account.idTokenClaims);

    Object.keys(tokenClaims).forEach((key) => {
        let row = tableBody.insertRow(0);
        let cell1 = row.insertCell(0);
        let cell2 = row.insertCell(1);
        let cell3 = row.insertCell(2);
        cell1.textContent = tokenClaims[key][0];
        cell2.textContent = tokenClaims[key][1];
        cell3.textContent = tokenClaims[key][2];
    });
};
