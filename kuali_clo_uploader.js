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

    exec(`npx electron ${tempPath}`, (error, stdout, stderr) => {
      fs.unlinkSync(tempPath);
      if (error) return reject(error);
      const selectedFile = stdout.trim();
      if (!selectedFile) return reject('No file selected');
      resolve(selectedFile);
    });
  });
}

async function launchBrowserOnly() {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    args: ['--start-maximized']
  });
  const page = await browser.newPage();
  await page.goto('https://york-sbx.kuali.co/cor/main/#/apps');
  console.log("🟢 Logged in? Please log in manually if prompted.");
  return { browser, page };
}

async function clickXPath(page, xpath, label) {
  await sleep(5000);
  const clicked = await page.evaluate((xp) => {
    const iterator = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    const el = iterator.singleNodeValue;
    if (el) {
      el.click();
      return true;
    }
    return false;
  }, xpath);
  console.log(clicked ? `✅ Clicked ${label}` : `⚠️ Could not click ${label}`);
  return clicked;
}

async function searchCourse(page, row) {
  const part1 = row.getCell('A').value;
  const part2 = row.getCell('C').value;
  const part3 = row.getCell('D').value;
  const courseCode = `${part1}/${part2} ${part3}`.replace(/\s+/g, ' ');

  await page.waitForSelector('#search-box', { timeout: 10000 });
  await page.evaluate(() => {
    document.querySelector('#search-box').value = '';
  });
  await page.type('#search-box', courseCode);
  console.log(`🔍 Searched for course: ${courseCode}`);
  await sleep(2000);
  await clickXPath(page, "//*[@id='0_']/a", "first course result");
}

async function inputCLO(page, cloText) {
  await sleep(1000);
  const success = await page.evaluate((text) => {
    const xpath = "//*[@id='61f570e185538c3f5de52c0b']/div/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div[1]/input";
    const iterator = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    const input = iterator.singleNodeValue;
    if (!input) return false;
    input.focus();
    input.value = text;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }, cloText);
  console.log(success ? "✅ Inputted CLO text" : "⚠️ Could not find CLO input box");
}

async function selectNativeDropdown(page, selector, valueToMatch) {
  const result = await page.evaluate((sel, val) => {
    const select = document.querySelector(sel);
    if (!select) return { success: false, reason: 'Dropdown not found', options: [] };

    const options = Array.from(select.options).map(opt => opt.textContent.trim());
    const matchedIndex = options.findIndex(o => o.toLowerCase().includes(val.toLowerCase()));

    if (matchedIndex === -1) return { success: false, val, options };

    select.selectedIndex = matchedIndex;
    select.dispatchEvent(new Event('input', { bubbles: true }));
    select.dispatchEvent(new Event('change', { bubbles: true }));

    return { success: true, selected: options[matchedIndex], options };
  }, selector, valueToMatch);

  console.log(`🎯 Native dropdown (${selector}):`, result);
  return result.success;
}

async function selectDropdownValues(page, gradAttrValue, gradAttrMapLevelValue) {
  const xpathGAI = '//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[2]/div/div[1]/div[2]/div/div[3]/div/select';
  const xpathGAML = '//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[2]/div/div[1]/div[2]/div/div[4]/div/select';

  const success = await page.evaluate((attrVal, mapLevelVal, xpGAI, xpGAML) => {
    function selectValue(xpath, desiredValue) {
      const iterator = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      const select = iterator.singleNodeValue;
      if (!select) return { success: false, reason: 'Dropdown not found', attempted: desiredValue };

      const options = Array.from(select.options).map(opt => opt.textContent.trim());
      const matchedIndex = options.findIndex(o => o.toLowerCase().includes(desiredValue.toLowerCase()));

      if (matchedIndex === -1) {
        return { success: false, reason: 'No match found', attempted: desiredValue, available: options };
      }

      select.selectedIndex = matchedIndex;
      select.dispatchEvent(new Event('input', { bubbles: true }));
      select.dispatchEvent(new Event('change', { bubbles: true }));

      return { success: true, selected: options[matchedIndex] };
    }

    const resultGAI = selectValue(xpGAI, attrVal);
    const resultGAML = selectValue(xpGAML, mapLevelVal);

    return {
      resultGAI,
      resultGAML,
      success: resultGAI.success && resultGAML.success
    };
  }, gradAttrValue, gradAttrMapLevelValue, xpathGAI, xpathGAML);

  console.log(`🔽 GAI Dropdown Result:`, success.resultGAI);
  console.log(`🔽 GAML Dropdown Result:`, success.resultGAML);
  console.log(success.success ? "✅ GAI & Mapping Level set" : "⚠️ Could not set GAI or Mapping Level");

  return success.success;
}


async function main() {
  try {
    const selectedFile = await askForFile();
    const resolvedPath = path.resolve(selectedFile);
    console.log("Excel file selected:", resolvedPath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(resolvedPath);
    const worksheet = workbook.worksheets[0];

    const { browser, page } = await launchBrowserOnly();

    console.log("➡️ After logging in, press ENTER here to continue...");
    process.stdin.resume();
    await new Promise(resolve => process.stdin.once('data', () => {
      process.stdin.pause();
      resolve();
    }));

    await clickXPath(page, "//div[contains(text(),'Curriculum')]", "Curriculum");
    await clickXPath(page, "//*[@id='app']/div/div[4]/nav/ul/li[3]/a/img", "Courses");

    const summary = [];

    const gamlMap = {
      "I": "Introduced",
      "D": "Developed",
      "A": "Applied/Used",
      "I,D": "Introduced & Developed",
      "I,A": "Introduced & Applied",
      "D,A": "Developed & Applied",
      "I,D,A": "Introduced, Developed & Applied"
    };

  let lastCourseCode = '';
  let cloBlockIndex = 1;
  let pendingTitles = [];

  for (let i = 2; i <= worksheet.actualRowCount; i++) {
    const row = worksheet.getRow(i);
    const cloText = row.getCell('G').value?.toString().trim() || '';
    const gradAttrValue = row.getCell('I').value?.toString().trim() || '';
    const gamlRaw = row.getCell('L').value?.toString().replace(/\s/g, '') || '';
    const faculty = row.getCell('A').value;
    const dept = row.getCell('C').value;
    const code = row.getCell('D').value;
    const courseCode = `${faculty}/${dept} ${code}`.replace(/\s+/g, ' ');
    const gradAttrMapLevelValue = gamlMap[gamlRaw] || '';
    const isSameCourse = courseCode === lastCourseCode;

    if (!isSameCourse && lastCourseCode !== '') {
      // === Phase 2: Input CLO titles for previous course ===
      for (const { text, divIndex } of pendingTitles) {
        const typed = await page.evaluate((text, divIndex) => {
          const xpath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${divIndex}]/div/div[1]/div[1]/div[1]/input`;
          const input = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
          if (!input) return false;
          input.focus();
          input.value = '';
          for (const char of text) {
            input.value += char;
            input.dispatchEvent(new InputEvent('input', { bubbles: true }));
          }
          input.dispatchEvent(new Event('change', { bubbles: true }));
          return true;
        }, text, divIndex);
        console.log(typed ? `✍️ CLO block ${divIndex} titled` : `⚠️ Could not title CLO block ${divIndex}`);
      }

      // === Save Course ===
      await clickXPath(page, '//*[@id="enrolmentNotes-input"]', 'Save (Enrolment Notes)');
      await sleep(500); // slight buffer
      await page.evaluate(() => {
        const input = document.evaluate('//*[@id="enrolmentNotes-input"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (input) {
          input.focus();
          input.value = ' ';
          input.dispatchEvent(new InputEvent('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
      await sleep(2000);

      // Reset state for next course
      pendingTitles = [];
      cloBlockIndex = 1;
    }

    if (!isSameCourse) {
      console.log(`\n📘 New course: ${courseCode}`);
      await searchCourse(page, row);
      await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[1]/div/div[1]/a", "Edit");
      await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[3]/nav/ul/li[10]/div", "Lassonde Course Outcomes");
    }

    const divIndex = cloBlockIndex + 1;
    const gaiXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${divIndex}]/div/div[1]/div[2]/div/div[3]/div/select`;
    const gamlXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${divIndex}]/div/div[1]/div[2]/div/div[4]/div/select`;

    await sleep(3000);
    await clickXPath(page, '//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[1]/div/div[2]/button[1]', "Add New CLO");
    await sleep(1000);

    const dropdownsOK = await page.evaluate((attrVal, mapVal, gaiXP, gamlXP) => {
      function setSelect(xpath, val) {
        const select = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!select) return false;
        const match = Array.from(select.options).find(opt =>
          opt.textContent?.toLowerCase().includes(val.toLowerCase())
        );
        if (!match) return false;
        select.value = match.value;
        select.dispatchEvent(new Event('input', { bubbles: true }));
        select.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      }
      return setSelect(gaiXP, attrVal) && setSelect(gamlXP, mapVal);
    }, gradAttrValue, gradAttrMapLevelValue, gaiXPath, gamlXPath);
    console.log(dropdownsOK ? "✅ GAI & GAML set" : "⚠️ Dropdown set failed");

    pendingTitles.push({ text: cloText, divIndex });
    lastCourseCode = courseCode;
    cloBlockIndex++;
    summary.push({ row: i, course: courseCode, status: '🛠 Pending Title' });
  }

  // === Final course save & title phase ===
  for (const { text, divIndex } of pendingTitles) {
    const typed = await page.evaluate((text, divIndex) => {
      const xpath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${divIndex}]/div/div[1]/div[1]/div[1]/input`;
      const input = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      if (!input) return false;
      input.focus();
      input.value = '';
      for (const char of text) {
        input.value += char;
        input.dispatchEvent(new InputEvent('input', { bubbles: true }));
      }
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }, text, divIndex);
    console.log(typed ? `✍️ CLO block ${divIndex} titled` : `⚠️ Failed to title CLO block ${divIndex}`);
  }

  // Final save
  await clickXPath(page, '//*[@id="enrolmentNotes-input"]', 'Save (Enrolment Notes)');
  await sleep(500);
  await page.evaluate(() => {
    const input = document.evaluate('//*[@id="enrolmentNotes-input"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (input) {
      input.focus();
      input.value = ' ';
      input.dispatchEvent(new InputEvent('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });
  await sleep(2000);




    console.log("\n📋 Summary:");
    summary.forEach(s => console.log(`Row ${s.row}: ${s.status}`));

    console.log("🟢 All done. Browser is still open for review or next steps.");
  } catch (err) {
    console.error("❌ Error:", err.message);
  }
}

main();
