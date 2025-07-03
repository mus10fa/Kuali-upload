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
  "2A": "02a - Formulate engineering problems",
  "2B": "02b - Solve engineering problems",
  "2C": "02c - Evaluate solutions to engineering problems",
  "3A": "03a - Demonstrate ability to plan the investigation of engineering problems",
  "3B": "03b - Demonstrate ability to collect data from experiments to investigate engineering problems",
  "3C": "03c - Demonstrate ability to synthesize data from experiments to investigate engineering problems",
  "4A": "04a - Identify requirements and specifications for complex, open-ended engineering design problems",
  "4B": "04b - Decompose complex systems into smaller, more manageable sub-systems",
  "4C": "04c - Develop and refine design solutions considering constraints including but not limited to health and safety risks, applicable standards, economic, environmental, cultural and societal considerations",
  "4D": "04d - Evaluate and compare engineering design solutions to advance to a final design",
  "5A": "05a - Select appropriate techniques, resources, and modern engineering tools",
  "5B": "05b - Apply appropriate techniques, resources, and modern engineering tools with identifications of limitations",
  "5C": "05c - Extend, adapt, or create appropriate techniques, resources, and modern engineering tools",
  "6A": "06a - Establish and review team organizations, goals, and responsibilities",
  "6B": "06b - Contribute as an active team member or leader to complete individual tasks",
  "6C": "06c - Collaborate with others to complete team goals effectively",
  "7A": "07a - Understand, interpret, and identify engineering knowledge from oral, written, or graphical communications",
  "7B": "07b - Orally present complex engineering concepts within the profession and to society at large",
  "7C": "07c - Produce written engineering reports and design documentation",
  "8A": "08a - Understand the role of the professional engineer in society, licensing, and the duty to protect the public interest",
  "8B": "08b - Describe the importance of standards, codes, regulations, best practices, laws, compliance with the Professional Engineers Act, and health and safety in engineering",
  "9A": "09a - Explain the relationship and impact of engineering projects on economic, health, safety, legal, and society issues and values",
  "9B": "09b - Evaluate the uncertainties and limitations in assessing the impact of engineering activities on economic, social, health, safety, legal, and cultural aspects of society",
  "10A": "10a - Explain ethical behaviour, value of decolonization, equity, diversity, and inclusion (DEDI), and importance of accountability in the workplace",
  "10B": "10b - Apply professional codes of ethics to engineering practices with respect to the role and responsibilities of individuals",
  "11A": "11a - Apply economic principles to support decision making and identify limitations of economics and business practices in engineering",
  "11B": "11b - Implement management process, build and monitor a project schedule based on tasks, milestones, and project risks, and make adjustments over time based on project status",
  "12A": "12a - Identify gaps in their knowledge, skills, and abilities",
  "12B": "12b - Obtain and evaluate information or training from appropriate sources",
  "12C": "12c - Develop goals and long term plans for continued learning to maintain professional standing and adapt to a changing world"
};

async function selectDropdownByText(page, xpath, matchText) {
  if (!matchText) {
    console.log(`⚠️ No match text provided for dropdown at: ${xpath}`);
    return false;
  }

  const selector = `xpath///${xpath.replace(/^\/+/, '')}`;
  const dropdown = await page.waitForSelector(selector).catch(() => null);
  if (!dropdown) {
    console.log(`⚠️ Could not find dropdown at: ${xpath}`);
    return false;
  }

  await page.evaluate((xp, mt) => {
    const select = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    const options = Array.from(select.options);
    const matchIndex = options.findIndex(opt => 
      typeof opt.textContent === 'string' && opt.textContent.trim().toLowerCase().includes(mt.toLowerCase())
    );
    if (matchIndex !== -1) {
      select.selectedIndex = matchIndex;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }, xpath, matchText);

  console.log(`✅ Selected "${matchText}"`);
  return true;
}

async function inputNewCLO(page, cloText = '', gaiText, gamlText, cloNumber) {
  const addCLOXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${cloNumber}]/div/div[2]/button[1]`;
  const cloInputXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${cloNumber + 1}]/div/div[1]/div[1]/div[1]/input`;
  const gaiXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${cloNumber + 1}]/div/div[1]/div[2]/div/div[3]/div/select`;
  const gamlXPath = `//*[@id="61f570e185538c3f5de52c0b"]/div/div/div/div/div[2]/div[${cloNumber + 1}]/div/div[1]/div[2]/div/div[4]/div/select`;

  const addBtn = await page.waitForSelector(`xpath///${addCLOXPath.replace(/^\/+/, '')}`);
  await addBtn.click();
  await sleep(1500);

  await selectDropdownByText(page, gaiXPath, gaiMap[gaiText?.toUpperCase()]);
  await selectDropdownByText(page, gamlXPath, gamlMap[gamlText]);

  const cloSelector = `xpath///${cloInputXPath.replace(/^\/+/, '')}`;
  const cloInput = await page.waitForSelector(cloSelector);
  await cloInput.click({ clickCount: 3 });
  await page.keyboard.press('Backspace');
  for (const char of cloText) {
    await page.keyboard.type(char, { delay: 75 });
  }

  await page.evaluate((xp) => {
    const input = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (input) {
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      input.dispatchEvent(new Event('blur', { bubbles: true }));
    }
  }, cloInputXPath);
  console.log(`✅ Typed CLO: ${cloText}`);
}

// The rest of the script remains the same
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
        console.log(`🔍 Preparing to input CLO for: ${courseCode}`);
        await searchCourse(page, courseCode);
        await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[1]/div/div[1]/a", "Edit");
        await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[3]/nav/ul/li[10]/div", "Lassonde Course Outcomes");
        lastCourseCode = courseCode;
      }
      
      await inputNewCLO(page, clo, gai, gaml, rowIndex - 1);
      console.log(`✅ CLO ${rowIndex - 1} input complete.`);
    }

    console.log("✅ All CLOs input complete. Ready for review.");
  } catch (err) {
    console.error("❌ Error in automation:", err);
  }
}

main();
