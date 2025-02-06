import test from "playwright/test";

test('login',async({page})=>{


 await page.goto("https://react-dev.gravitywrite.com/login");

 const emailLocator = page.locator("#text");
 await emailLocator.click();
 await emailLocator.fill("jerald@wl.team");
 const passwordLocator = page.locator("#password");
 await passwordLocator.click();
 await passwordLocator.fill("Test@123");
 const button = page.locator("//button[@type='submit']"); 
 await button.click();
 const loginOtherPageInfo = page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();
                if (await loginOtherPageInfo) {
                    const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                    await continueButton.click();
                }
                const [response] = await Promise.all([
                  page.waitForResponse(
                    (res) =>
                      res.status() == 200 &&
                      res.url() === 'https://dev.gravitywrite.com/api/auth/login'
                  ),
                ]);
                const responseData = await response.json();
                console.log(responseData);
                console.log('User Subscription:', responseData.data.user.user_subscription.plan.title);



})

