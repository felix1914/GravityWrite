import { test } from "playwright/test";

test('login', async ({ page }) => {
    // Navigate to login page
    await page.goto("https://react-dev.gravitywrite.com/login");

    // Perform login actions
    const emailLocator = page.locator("#text");
    await emailLocator.fill("jerald@wl.team");

    const passwordLocator = page.locator("#password");
    await passwordLocator.fill("Test@123");

    const button = page.locator("//button[@type='submit']");
    await button.click();

    // Wait for navigation or page response after login
  
    // Declare the 'total' variable in a higher scope
    let total = 0;

    // Check if redirected or new page content appears
    const loginOtherPageInfo = page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();

    if (await loginOtherPageInfo) {
        const continueButton = page.locator("//button[contains(text(),'Continue here')]");
        await continueButton.click();

        const [loginResponse] = await Promise.all([
            page.waitForResponse(
                (res) =>
                    res.status() === 200 &&
                    res.url() === 'https://dev.gravitywrite.com/api/auth/login'
            ),
        ]);

        const responseData = await loginResponse.json();
        total = responseData.data.user.user_subscription.plan.word_limit;
        console.log('Total Word Limit:', total);

        // Wait for navigation or content load after clicking continue
    }

    const offerButton = page.locator("//p[contains(text(),'No thanks!')]");
    await offerButton.click();

    const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
    await skipSurveyButton.click();

    const tryNowButton = page.locator("(//p[contains(text(),'Not Now')])[2]");
    await tryNowButton.click();

    // Navigate to the new page and interact with it
    await page.locator("//span[contains(text(),'Book Details Generator for Amazon')]").click();

    // Ensure the page has fully loaded
    await page.waitForSelector('textarea[name="question_490"]');

    const input = page.locator('textarea[name="question_490"]');
    await input.fill("Book");

    const createButton = page.locator("//button[contains(text(),'Create Content')]");
    await createButton.click();

    // Wait for API response
    const [storeContentResponse] = await Promise.all([
        page.waitForResponse(
            (res) =>
                res.status() === 200 &&
                res.url() === 'https://dev.gravitywrite.com/api/singlePrompt/storeContent'
        ),
    ]);

    const storeContentData = await storeContentResponse.json();
    const usedWords = storeContentData.data.usage;

    console.log('Response Data:', storeContentData);
    console.log('Used Words:', usedWords);

    // Compare the total word limit with used words
    if (total < usedWords) {
        console.log("Yes");
    } else {
        console.log("No");
    }
});
