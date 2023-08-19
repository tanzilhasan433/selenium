


const webdriver = require('selenium-webdriver');
const { Builder, By, Key, until } = webdriver;
const chrome = require('selenium-webdriver/chrome');
const exceljs = require('exceljs');

    // Set up WebDriver
    const driver = new Builder()
    .forBrowser('chrome')
    .setChromeOptions(new chrome.Options().headless())
    .build();

        async function getSuggestions(searchQuery) {
        // Open the search engine
        await driver.get('https://www.google.com');

        // Enter search query and get suggestions
        const searchBox = await driver.findElement(By.name('q'));

        await searchBox.sendKeys(searchQuery, Key.RETURN);

        // Wait for suggestions to load (you might need to adjust the wait time)
        await driver.wait(until.elementLocated(By.css('ul.G43f7e li')), 5000);

        // Get suggestion elements
        const suggestionElements = await driver.findElements(By.css('ul.G43f7e li'));

        const suggestions = await Promise.all(suggestionElements.map(element => element.getText()));
        const allSuggestions = suggestions.map(suggestion => suggestion.split('\n')[0]);

        const sortedSuggestions = allSuggestions.sort((a, b) => a.length - b.length);
        const shortest = sortedSuggestions[0];
        const largest = sortedSuggestions[sortedSuggestions.length - 1];

        return [largest, shortest];
        }

        async function main() {
        const excelFilePath = 'Excel.xlsx';
        const workbook = new exceljs.Workbook();

        await workbook.xlsx.readFile(excelFilePath);
        const sheetNames = workbook.sheetNames;

        for (const sheetName of sheetNames) {
            const worksheet = workbook.getWorksheet(sheetName);

            const thirdColumn = worksheet.getColumn(3).slice(1);

            for (let i = 0; i < thirdColumn.length; i++) {
            const each = thirdColumn[i].value;
            const result = await getSuggestions(each);
            worksheet.getCell(`D${i + 2}`).value = result[0];
            worksheet.getCell(`E${i + 2}`).value = result[1];
            }

            await workbook.xlsx.writeFile(excelFilePath);
        }

        await driver.quit();
        }

     main().catch(error => console.error(error));

        