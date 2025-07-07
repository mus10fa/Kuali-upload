const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');
const ExcelJS = require('exceljs');

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function askForFile() {
  return new Promise((resolve, reject) => {
    const electronCode = `
      const { app, dialog } = require('electron');
      app.whenReady().then(() => {
        dialog.showOpenDialog({
          title: 'Select Excel File',
          filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xlsm'] }],
          properties: ['openFile']
        }).then(result => {
          if (!result.canceled) {
            console.log(result.filePaths[0]);
          } else {
            console.log('');
          }
          app.exit();
        });
      });
    `;
    const tempPath = path.join(__dirname, 'file_chooser_temp.js');
    fs.writeFileSync(tempPath, electronCode);

    exec(`npx electron ${tempPath}`, (error, stdout) => {
      fs.unlinkSync(tempPath);
      if (error) return reject(error);
      const selectedFile = stdout.trim();
      if (!selectedFile) return reject('No file selected');
      resolve(selectedFile);
    });
  });
}

const gamlMap = {
  "I": "Introduced",
  "D": "Developed",
  "A": "Applied/Used",
  "I,D": "Introduced & Developed",
  "D,A": "Developed & Applied",
  "I,A": "Introduced & Applied",
  "I,D,A": "Introduced, Developed & Applied"
};

const gaiMap = {
  "1A": "01a - Demonstrate competence in mathematics",
  "1B": "01b - Demonstrate foundational knowledge of natural sciences",
  "1C": "01c - Demonstrate knowledge of engineering fundamentals",
  "1D": "01d - Demonstrate competence in specialized engineering knowledge",
  // Add the rest as needed...
};

async function selectDropdownByLabelText(page, container, labelText, visibleOptionText) {
  if (!visibleOptionText) {
    console.log(`⚠️ No value provided to select for label "${labelText}"`);
    return;
  }
  const dropdownHandle = await page.evaluateHandle((container, labelText) => {
    const labelEls = container.querySelectorAll('label');
    for (const label of labelEls) {
      if (label.innerText.trim().toLowerCase().includes(labelText.toLowerCase())) {
        const parent = label.closest('div');
        if (!parent) continue;
        const select = parent.querySelector('select');
        return select || null;
      }
    }
    return null;
  }, container, labelText);

  if (!dropdownHandle) {
    console.log(`⚠️ Dropdown with label '${labelText}' not found`);
    return;
  }

  await page.evaluate((select, visibleText) => {
    const options = Array.from(select.options);
    const match = options.find(o =>
      o.textContent && o.textContent.trim().toLowerCase().includes(visibleText.toLowerCase())
    );
    if (match) {
      select.value = match.value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }, dropdownHandle, visibleOptionText);

  await dropdownHandle.dispose();
}

async function inputNewCLO(page, cloText = '', gaiText, gamlText) {
  const addButtons = await page.$x("//button[@aria-label='Add outcome']");
  if (addButtons.length === 0) throw new Error("No 'Add outcome' buttons found");
  const addBtn = addButtons[addButtons.length - 1];

  await addBtn.click();
  await sleep(1500);
  console.log('➡️ Clicked Add button. Now trying to input CLO and select dropdowns...');

  const containerHandle = await addBtn.evaluateHandle(btn => btn.closest('div[style*="flex"]'));
  if (!containerHandle) throw new Error("Could not find container div near Add button");

  const cloInput = await containerHandle.asElement().$('input.form-control');
  if (!cloInput) throw new Error("CLO input not found");
  await cloInput.click({ clickCount: 3 });
  await page.keyboard.press('Backspace');
  await cloInput.type(cloText, { delay: 75 });
  console.log(`✅ CLO input: "${cloText}"`);

  console.log('📌 GAI raw:', gaiText);
  console.log('📌 GAI mapped:', gaiMap[gaiText?.toUpperCase()]);
  console.log('📌 GAML raw:', gamlText);
  console.log('📌 GAML mapped:', gamlMap[gamlText]);

  await selectDropdownByLabelText(page, containerHandle, 'Graduate Attribute Indicator', gaiMap[gaiText?.toUpperCase()]);
  await selectDropdownByLabelText(page, containerHandle, 'Graduate Attribute Map Level', gamlMap[gamlText]);

  await containerHandle.dispose();
  console.log(`✅ Finished inputting CLO: ${cloText}`);
}

async function launchBrowserOnly() {
  const browser = await puppeteer.launch({ headless: false, defaultViewport: null, args: ['--start-maximized'] });
  const page = await browser.newPage();
  await page.goto('https://york-sbx.kuali.co/cor/main/#/apps');
  console.log("🟢 Logged in? Please log in manually if prompted.");
  return { browser, page };
}

async function navigateToCourses(page) {
  const curriculumXPath = '//*[@id="mainContent"]/div[1]/div/ul/li[1]/div[2]';
  await clickXPath(page, curriculumXPath, 'Curriculum');
  await sleep(3000);
  const coursesXPath = "//*[@id='app']/div/div[4]/nav/ul/li[3]/a/img";
  await clickXPath(page, coursesXPath, "Courses");
}

async function searchCourse(page, courseCode) {
  await page.waitForSelector('#search-box', { timeout: 10000 });
  await page.evaluate(() => { document.querySelector('#search-box').value = ''; });
  await page.type('#search-box', courseCode);
  console.log(`🔍 Searched for course: ${courseCode}`);
  await sleep(2000);
  const firstResultXPath = "//*[@id='0_']/a";
  await clickXPath(page, firstResultXPath, "first course result");
}

async function clickXPath(page, xpath, label) {
  const selector = `xpath///${xpath.replace(/^\/+/, '')}`;
  const element = await page.waitForSelector(selector);
  await element.click();
  console.log(`✅ Clicked ${label}`);
}

async function main() {
  try {
    const selectedFile = await askForFile();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(selectedFile);
    const sheet = workbook.worksheets[0];

    const { browser, page } = await launchBrowserOnly();
    await page.waitForFunction(() => !location.href.includes('passportyork'), { timeout: 0 });
    console.log("✅ Login complete. Waiting for page to load...");
    await sleep(60000);

    await navigateToCourses(page);

    let lastCourseCode = null;

    for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex++) {
      const row = sheet.getRow(rowIndex);
      const faculty = row.getCell('A').value;
      const dept = row.getCell('C').value;
      const code = row.getCell('D').value;
      const clo = row.getCell('G').value;
      const gai = row.getCell('I').value;
      const gaml = row.getCell('L').value;
      const courseCode = `${faculty}/${dept} ${code}`.replace(/\s+/g, ' ');

      if (courseCode !== lastCourseCode) {
        if (lastCourseCode) await sleep(3000);
        await clickXPath(page, "//*[@id='app']/div/div[4]/nav/ul/li[3]/a/img", "Courses");
        console.log(`🔍 Preparing to input CLO for: ${courseCode}`);
        await searchCourse(page, courseCode);
        await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[1]/div/div[1]/a", "Edit");
        await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[3]/nav/ul/li[10]/div", "Lassonde Course Outcomes");
        lastCourseCode = courseCode;
      }

      await inputNewCLO(page, clo, gai, gaml);
    }

    console.log("✅ All CLOs input complete. Ready for review.");
  } catch (err) {
    console.error("❌ Error in automation:", err);
  }
}

main();
