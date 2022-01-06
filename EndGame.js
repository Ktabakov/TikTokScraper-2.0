const { chromium } = require(`playwright-chromium`);
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
const fs = require('fs');

const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./SpainTikTok.xlsx');
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

        //Reset context each 300-500 cycles
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
            //wait for crash -- adjust div id
            await page.waitForSelector('div')
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
                //first replace "window['SIGI_STATE']=" with "" -- .replace
                //second replace "" with "" -- substring -- s = s.substring(0, s.indexOf("; however, if these thoughts are taking up"));
                //add end symbols "\"}}}}}"
                //const element = document.getElementById('__NEXT_DATA__')               
                let element = document.getElementById('sigi-persisted-data')
                let q = element.textContent.replace("window['SIGI_STATE']=", "")
                let w = q.substring(0, q.indexOf("; however, if these thoughts are taking up"))
                let final = w + "\"}}}}}"

                let json = JSON.parse(final)
                let itemModule = json["ItemModule"]
                let userModuleUsers = json["UserModule"]["users"]
                const uniqueID = Object.keys(itemModule)[0]
                const userUniqueName = Object.keys(userModuleUsers)[0]

                let author = json["ItemModule"][uniqueID]["author"]
                console.log(author)
                let id = json["ItemModule"][uniqueID]["id"]
                let dateObject = new Date(json["ItemModule"][uniqueID]["createTime"] * 1000)
                let createDate = dateObject.toLocaleDateString("uk-Uk");
                let description = json["ItemModule"][uniqueID]["desc"]
                let verified = json["UserModule"]["users"][userUniqueName]["verified"]
                let video = json["ItemModule"][uniqueID]["video"]
                let music = json["ItemModule"][uniqueID]["music"]
                let songTitle = json["ItemModule"][uniqueID]["music"]["title"]
                let authorName = json["ItemModule"][uniqueID]["music"]["authorName"]
                let musicURL = json["ItemModule"][uniqueID]["music"]["playUrl"]


                return { music, id, description, author, video, createDate, songTitle, authorName, verified, musicURL};
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
        "musicURL"
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
