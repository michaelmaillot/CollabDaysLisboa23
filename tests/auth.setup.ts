import { test as setup, expect } from '@playwright/test';

const authFile = 'playwright/.auth/user.json';

setup('authenticate', async ({ page, context }) => {
  const login = process.env.M365_LOGIN;
  const pwd = process.env.M365_PWD;

  if (login && pwd) {
    await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=4765445b-32c6-49b0-83e6-1d93765276ca&redirect_uri=https%3A%2F%2Fwww.office.com%2F&response_type=code%20id_token&scope=openid%20profile%20https%3A%2F%2Fwww.office.com%2Fv2%2FOfficeHome.All');
    const loginPage = page.locator('#i0116');
    await loginPage.isVisible();
    await loginPage.fill(login);

    const validate = page.locator('#idSIButton9');
    await validate.click();
    await page.locator('#i0118').fill(pwd);
    await validate.click();

    // 👇 Test can be made to be sure auth is successful
    // await page.waitForResponse("https://login.microsoftonline.com/common/login");
    // await validate.click();
    // await page.goto('https://microsoft365.com/?auth=2');
    // await expect(page.getByRole('heading', { name: 'Welcome to Microsoft 365' })).toBeVisible({timeout: 20000});

    await page.context().storageState({ path: process.env.CI ? 'storage.json' : authFile });
  }
  else {
    throw new Error('M365_LOGIN and M365_PWD environment variables must be set');
  }

  await context.close();
});