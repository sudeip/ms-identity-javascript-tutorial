/**
 * @jest-environment jsdom
 */

const request = require('supertest');
const path = require('path');
const fs = require('fs');
const { TextDecoder, TextEncoder } = require('util');

Object.assign(global, { TextDecoder, TextEncoder });
const app = require('./server.js');

jest.dontMock('fs');

const html = fs.readFileSync(path.resolve(__dirname, './public/index.html'), 'utf8');

describe('Sanitize index page', () => {
    beforeAll(async() => {
        global.document.documentElement.innerHTML = html.toString();
    });

    it('should NOT load msal-browser — a spoke app never talks to Entra ID directly', () => {
        expect(document.getElementById('load-msal')).toBeNull();
    });
});

describe('Sanitize spoke configuration', () => {
    beforeAll(() => {
        global.AUTHWEB_HUB_URL = require('./public/authConfig.js').AUTHWEB_HUB_URL;
    });

    it('should point at AuthWeb Hub, not embed any Entra ID credentials', () => {
        expect(AUTHWEB_HUB_URL).toBeDefined();
        expect(AUTHWEB_HUB_URL).toMatch(/^https?:\/\//);
    });
});

describe('Ensure pages served', () => {

    beforeAll(() => {
        process.env.NODE_ENV = 'test';
    });

    it('should get index page', async () => {
        const res = await request(app)
            .get('/');

        const data = await fs.promises.readFile(path.join(__dirname, './public/index.html'), 'utf8');
        expect(res.statusCode).toEqual(200);
        expect(res.text).toEqual(data);
    });
});
