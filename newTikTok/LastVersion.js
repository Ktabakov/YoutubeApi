const { chromium } = require(`playwright-chromium`);
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
const fs = require('fs');

const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./Links.xlsx');
const sheets = spreadsheet.SheetNames;
const firstSheet = spreadsheet.Sheets[sheets[0]]; //sheet 1 is index 0

(async () => {

    let links = [];

    for (let i = 1; ; i++) {
        const firstColumn = firstSheet['A' + i];
        if (!firstColumn) {
            break;
        }
        links.push(firstColumn.h);
    }
    let items = [];
    const browser = await chromium.launch({ headless: false });
    let context = await browser.newContext()

    let page = await context.newPage();

    await page.setDefaultNavigationTimeout(0);

    await page.goto('https://www.tiktok.com/');
    await page.waitForTimeout(10000)
    let cookies = await context.cookies()
    let cookieJson = JSON.stringify(cookies)
    fs.writeFileSync('cookies.json', cookieJson)

    let i = 0;
    let item;

    for (let link of links) {
        console.log("Fetch", link)

        //Reset context each 500-600 cycles
        //use 3 to test
        if (i % 300 == 0 && i != 0) {
            await context.close();
            await page.close();
            context = await browser.newContext()
            page = await context.newPage();
            await page.setDefaultNavigationTimeout(0);

            cookies = fs.readFileSync('cookies.json', 'utf8')
            let deserializedCookies = JSON.parse(cookies)
            await context.addCookies(deserializedCookies)
        }

        try {
            await page.goto(link)
            await page.waitForSelector('div[id=__next]')
            i++;
        }
        catch (error) {
            item = {
                link: link,
                error: item,
            }
            console.log(item)
            items.push(item)
            console.log("Died In First Catch")
            i++;
            continue;
        }

        item = await page.evaluate(() => {
            try {
                const element = document.getElementById('__NEXT_DATA__')
                const json = JSON.parse(element.textContent)
                const id = json["props"]["pageProps"]["feedConfig"]["id"]
                const description = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["desc"]
                const dateObject = new Date(json["props"]["pageProps"]["itemInfo"]["itemStruct"]["createTime"] * 1000)
                const createDate = dateObject.toLocaleDateString("uk-Uk");
                const video = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["video"]
                const author = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["author"]
                const music = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["music"]
                const songTitle = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["music"]["title"]
                const authorName = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["music"]["authorName"]
                const verified = json["props"]["pageProps"]["itemInfo"]["itemStruct"]["author"]["verified"]

                return { music, id, description, author, video, createDate, songTitle, authorName, verified };
            } catch (error) {
                return error
            }
        })

        item.link = link
        // If id is not set, it should be an error => log it
        if (!item.id) {
            item = {
                link: link,
                error: item,
            }
            console.log(item)
        }
        items.push(item)
    }

    const outputFields = [
        "id",
        "description",
        "createDate",
        "video",
        "author",
        "music",
        "songTitle",
        "authorName",
        "link",
        "verified",
    ]

    for (let i = 0; i < outputFields.length; i++) {
        worksheet.cell(1, i + 1).string(outputFields[i])
    }

    for (let index = 0; index < items.length; index++) {
        let item = items[index];
        if (item.id) {
            for (let i = 0; i < outputFields.length; i++) {
                let value = item[outputFields[i]];
                if (typeof value != "string") {
                    value = JSON.stringify(value)
                }
                worksheet.cell(index + 2, i + 1).string(value)
            }
        } else {
            worksheet.cell(index + 2, 2).string(item.link)
            worksheet.cell(index + 2, 3).string(item.error)
        }
    }

    workbook.write('Props.xlsx')
    console.log('Done!')

    await browser.close()
})();

console.log("non-async finished")
