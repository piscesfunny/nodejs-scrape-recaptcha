const puppeteer = require('puppeteer');
const request = require('request-promise-native')
const poll = require('promise-poller').default
const path = require('path');

const { getDataFromExcel, writeUniqueNumbersToExcel, writeResultToExcel, getUniqueArray } = require('./excel')

const config = {
    sitekey: '6LdpYRoTAAAAAAL0_J6lZ7LIKHD7bX6T2_Rgd-UB',
    pageurl: 'https://ownvehicle.askmid.com',
    apiKey: 'f38cbfee4096340a79eb5ea447dece11',
}

const chromeOptions = {
    headless:true,
    defaultViewport: null,
    slowMo:100,
};

const dataDirectoryPath = path.join(__dirname, './data')
const sourceExcelPath = path.join(dataDirectoryPath, 'origin_source.xls')

let insuredNumbers = []
let unInsuredNumbers = []
let resultText = ''

const getRegistrationNumbers = () => {
    const rawRegistrationNumbers = getDataFromExcel(sourceExcelPath)
    const registrationNumbers = getUniqueArray(rawRegistrationNumbers)

    const resultExcelPath = path.join(dataDirectoryPath, 'filtered_source.xlsx')

    writeUniqueNumbersToExcel(registrationNumbers, resultExcelPath)

    return registrationNumbers
}

async function initiateCaptchaRequest(apiKey) {
    const formData = {
        method: 'userrecaptcha',
        googlekey: config.sitekey,
        key: apiKey,
        pageurl: config.pageurl,
        json: 1
    };
    const response = await request.post('http://2captcha.com/in.php', {form: formData})
    return JSON.parse(response).request
}

async function pollForRequestResults(key, id, retries = 50, interval = 1500, delay = 15000) {
    await timeout(delay)
    return poll({
        taskFn: requestCaptchaResults(key, id),
        interval,
        retries
    })
}

function requestCaptchaResults(apiKey, requestId) {
    const url = `http://2captcha.com/res.php?key=${apiKey}&action=get&id=${requestId}&json=1`
    return async function() {
        return new Promise(async function(resolve, reject){
            const rawResponse = await request.get(url);
            const resp = JSON.parse(rawResponse);
            if (resp.status === 0) return reject(resp.request);
            resolve(resp.request);
        })
    }
}

const timeout = millis => new Promise(resolve => setTimeout(resolve, millis))

const writeFinalResult = () => {
    const resultFilePath = path.join(dataDirectoryPath, 'result.xlsx')
    writeResultToExcel(insuredNumbers, unInsuredNumbers, resultFilePath)
}

async function doScraping(startIndex, countPerParallelExecution, registrationNumbers) {
    console.log(`${startIndex} : ${countPerParallelExecution}`)

    const browser = await puppeteer.launch(chromeOptions);
    const page = await browser.newPage();

    for(let x = startIndex; x < startIndex + countPerParallelExecution; x++) {
        await page.goto('https://ownvehicle.askmid.com',  {
            timeout: 0,
            waitUntil: "networkidle0"
        })

        const registrationNumber = registrationNumbers[x]

        if (registrationNumber === undefined) {
            break
        }

        await page.click('#acceptCookieBtn')
        await page.type('#txtVRN', registrationNumber)
        await page.click('#chkDataProtection')
        try {
            const requestId = await initiateCaptchaRequest(config.apiKey)
            const response = await pollForRequestResults(config.apiKey, requestId);
            await page.evaluate(`document.getElementById("g-recaptcha-response").innerHTML="${response}";`)
        } catch (e) {
            console.log(`Getting reCAPTCHA response from 2captcha.com failed - Registration Number : ${registrationNumber}`)
            continue
        }

        await timeout(500)

        await page.click('#btnCheckVehicle')
        await timeout(3000)

        const isByPassedCaptcha = await page.evaluate(() => {
            let isByPassedCaptcha = false
            try {
                document.querySelector('#captchaError').innerHTML
            } catch (e) {
                isByPassedCaptcha = true
            }

            return isByPassedCaptcha
        })

        if (!isByPassedCaptcha) {
            console.log(`Bypassing captcha failed for the Registration Number - ${registrationNumber}`)
            continue
        }

        await page.waitForSelector('#leftDiv > div > div > h4')

        resultText = await page.evaluate(() => {
            return document.querySelector('#leftDiv > div > div > h4 > b').innerText.trim();
        })

        console.log(`${resultText} - ${startIndex} : ${registrationNumber}`)
        if (resultText === 'YES.') insuredNumbers.push(registrationNumber)
        else  unInsuredNumbers.push(registrationNumber)

        await timeout(250)
    }

    await browser.close()
}

async function main() {
    const registrationNumbers = getRegistrationNumbers()
    // const totalCount = registrationNumbers.length
    const totalCount = 10
    const parallelExecutionCount = 5

    const countPerParallelExecution = Math.ceil(totalCount / parallelExecutionCount)
    let  startIndexes = []
    let startIndex = 0
    while (true) {
        if (startIndex > totalCount - 1) {
            break
        }

        startIndexes.push(startIndex)
        startIndex += countPerParallelExecution
    }

    await Promise.all(
        startIndexes.map(async index => {
            await doScraping(index, countPerParallelExecution, registrationNumbers)
        })
    )

    // await doScraping(0, 60, registrationNumbers)

    writeFinalResult()
}

main()
    .then(res=> console.log('Finished !!!'))
    .catch(err => console.error(err))