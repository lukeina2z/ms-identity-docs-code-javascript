/**
 * @jest-environment jsdom
 */

const { TextEncoder, TextDecoder } = require('util');

Object.assign(global, { TextEncoder, TextDecoder });

const request = require('supertest');
const path = require('path');
const fs = require('fs');

const app = require('./server.js');

jest.dontMock('fs');

const publicDir = path.join(__dirname, 'public');
const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');

describe('Sanitize index page', () => {
    beforeAll(async() => {
        global.document.documentElement.innerHTML = html.toString();
    });

    it('should load the locally served MSAL browser script', () => {
        expect(document.querySelector('script[src="/msal-browser.min.js"]')).not.toBeNull();
    });
});

describe('Sanitize configuration object', () => {
    beforeAll(() => {
        global.msalConfig = require('./public/authConfig.js').msalConfig;
    });

    it('should define the config object', () => {
        expect(msalConfig).toBeDefined();
    });

    it('should not contain credentials', () => {
        expect(msalConfig.auth.clientSecret).toBeUndefined();
        expect(msalConfig.auth.clientAssertion).toBeUndefined();
    });

    it('should contain authority URI', () => {
        const regexUri = /[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)?/gi;
        expect(regexUri.test(msalConfig.auth.authority)).toBe(true);
    });
});

describe('Ensure pages served', () => {

    it('should get index page', async () => {
        const res = await request(app)
            .get('/');

        const data = await fs.promises.readFile(path.join(publicDir, 'index.html'), 'utf8');
        expect(res.statusCode).toEqual(200);
        expect(res.text).toEqual(data);
    });

    it('should get signout page', async () => {
        const res = await request(app)
            .get('/signout');

        const data = await fs.promises.readFile(path.join(publicDir, 'signout.html'), 'utf8');
        expect(res.statusCode).toEqual(200);
        expect(res.text).toEqual(data);
    });

    it('should get redirect page', async () => {
        const res = await request(app)
            .get('/redirect');

        const data = await fs.promises.readFile(path.join(publicDir, 'redirect.html'), 'utf8');
        expect(res.statusCode).toEqual(200);
        expect(res.text).toEqual(data);
    });

    it('should get the MSAL browser script', async () => {
        const res = await request(app)
            .get('/msal-browser.min.js');

        const data = await fs.promises.readFile(path.join(__dirname, 'node_modules', '@azure', 'msal-browser', 'lib', 'msal-browser.min.js'), 'utf8');
        expect(res.statusCode).toEqual(200);
        expect(res.text).toEqual(data);
    });
});