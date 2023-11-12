import { test as setup, expect } from '@playwright/test';

const authFile = 'playwright/.auth/user.json';

setup('authenticate', async ({ page, context }) => {
    await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=4765445b-32c6-49b0-83e6-1d93765276ca&redirect_uri=https%3A%2F%2Fwww.office.com%2Flandingv2&response_type=code%20id_token&scope=openid%20profile%20https%3A%2F%2Fwww.office.com%2Fv2%2FOfficeHome.All&sso_reload=true');
    const login = page.locator('#i0116');
    await login.isVisible();
    await login.fill(process.env.M365_LOGIN);

    const validate = page.locator('#idSIButton9');
    await validate.click();
    await page.locator('#i0118').fill(process.env.M365_PWD);
    await validate.click();

    await page.waitForResponse("https://login.microsoftonline.com/common/login?sso_reload=true");
    await validate.click();
    await page.goto('https://microsoft365.com/?auth=2');
    await expect(page.getByRole('heading', { name: 'Welcome to Microsoft 365' })).toBeVisible();

    await page.context().storageState({ path: authFile });

    await context.close();
});