## IQTrackIT NG Playwright Tests

JavaScript end-to-end tests for the IQTrackIT NG portal at `https://iqtrackitng-test.azurewebsites.net/anmeldung`, using Playwright.

### Prerequisites

- Node.js and npm installed on your machine.

### Setup

```bash
cd iqtrackitng-tests
npm install
npx playwright install
```

This installs `@playwright/test` (see `package.json`) and the Chromium browser binaries.

### Running the tests

Run the full suite - headless

```bash
npm test
```

Run tests in headed mode:

```bash
npm run test:headed
```

Open the Playwright UI:

```bash
npm run test:ui
```

### Current coverage

- Login as **provider** using:
  - Email: `abad@mailinator.com`
  - Password: `test1234`
- Navigate to the **Meter Dashboard** menu after login.

Selectors for the login page are based on the visible English texts (`Enter Email or Mobile Number`, `Password`, `Sign In`) from the live site. If your app changes labels or you use German by default, you may need to tweak the locators in `tests/login-meter-dashboard.spec.js`.
