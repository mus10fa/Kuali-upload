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
  "1E": "01e - Demonstrate skills in programming, testing, and communication",
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

// Add this function to delete all existing CLOs
async function deleteAllExistingCLOs(page) {
  console.log("🗑️ Starting to delete all existing CLOs...");
  
  try {
    // Check if we need to delete anything
    const addNewSpans = await page.$x("//span[contains(text(),'Add New')]");
    if (addNewSpans.length > 0) {
      console.log("✅ No existing CLOs to delete - 'Add New' button is already visible");
      return;
    }
    
    // Keep deleting until the "Add New" button appears
    let deletionCount = 0;
    let continueDeleting = true;
    
    while (continueDeleting) {
      // Find all delete buttons with the trash icon
      const deleteButtonXPath = "//button[.//i[contains(@class, 'fa-trash')]]";
      const deleteButtons = await page.$x(deleteButtonXPath);
      
      if (deleteButtons.length === 0) {
        console.log("❓ No more delete buttons found, but 'Add New' not visible yet");
        break;
      }
      
      // Click the last delete button
      const lastDeleteButton = deleteButtons[deleteButtons.length - 1];
      await lastDeleteButton.click();
      deletionCount++;
      console.log(`🗑️ Deleted CLO #${deletionCount}`);
      
      // Wait for deletion to complete
      await sleep(1);
      
      // Check if "Add New" button has appeared
      const newAddNewSpans = await page.$x("//span[contains(text(),'Add New')]");
      if (newAddNewSpans.length > 0) {
        console.log("✅ 'Add New' button is now visible. All CLOs deleted.");
        continueDeleting = false;
      }
      
      // Safety check to prevent infinite loops
      if (deletionCount > 100) {
        console.log("⚠️ Safety limit reached (100 deletions). Breaking loop.");
        break;
      }
    }
    
    console.log(`✅ Finished deleting ${deletionCount} CLOs`);
  } catch (error) {
    console.error("❌ Error while deleting CLOs:", error);
  }
}

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
  const addNewSpans = await page.$x("//span[contains(text(),'Add New')]");
  if (addNewSpans.length > 0) {
    await addNewSpans[addNewSpans.length - 1].click();
    await page.waitForXPath("//button[@aria-label='Add outcome']", { timeout: 3000 });
  }

  let addButtons = await page.$x("//button[@aria-label='Add outcome']");
  if (addButtons.length === 0) throw new Error("No 'Add outcome' buttons found");
  const addBtn = addButtons[addButtons.length - 1];

  await addBtn.click();
  await page.waitForSelector('input.form-control', { timeout: 3000 });
  console.log('➡️ Clicked Add button. Now trying to input CLO and select dropdowns...');

  const containerHandle = await addBtn.evaluateHandle(btn => btn.closest('div[style*="flex"]'));
  if (!containerHandle) throw new Error("Could not find container div near Add button");

  const cloInput = await containerHandle.asElement().$('input.form-control');
  if (!cloInput) throw new Error("CLO input not found");

  // Fast paste input value
  await page.evaluate((input, value) => {
    input.value = '';
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }, cloInput, cloText);

  // Simulate space keypress to trigger key-related listeners
  await cloInput.focus();
  await page.keyboard.press('Space');

  console.log(`✅ CLO input pasted: "${cloText}"`);

  let gaiRaw = String(gaiText || '').trim().toUpperCase();

  console.log('📌 GAI raw:', gaiText);
  console.log('📌 GAI final:', gaiRaw);
  console.log('📌 GAI mapped:', gaiMap[gaiRaw]);
  console.log('📌 GAML raw:', gamlText);
  console.log('📌 GAML mapped:', gamlMap[gamlText]);

  await selectDropdownByLabelText(page, containerHandle, 'Graduate Attribute Indicator', gaiMap[gaiRaw]);
  await selectDropdownByLabelText(page, containerHandle, 'Graduate Attribute Map Level', gamlMap[gamlText]);

  await containerHandle.dispose();
  console.log(`✅ Finished inputting CLO: ${cloText}`);
}

async function clickXPath(page, xpath, label) {
  const selector = `xpath///${xpath.replace(/^\/+/g, '')}`;
  const element = await page.waitForSelector(selector, { timeout: 60000 });
  await element.click();
  console.log(`✅ Clicked ${label}`);
}

async function navigateToCourses(page) {
  const curriculumXPath = '//*[@id="mainContent"]/div[1]/div/ul/li[1]/div[2]';
  await clickXPath(page, curriculumXPath, 'Curriculum');
  await sleep(1000);
  const coursesXPath = "//*[@id='app']/div/div[4]/nav/ul/li[3]/a/img";
  await clickXPath(page, coursesXPath, "Courses");
}

async function searchCourse(page, courseCode) {
  if (!courseCode || courseCode.includes("null")) {
    console.log("⚠️ Invalid or missing course code. Skipping...");
    return false;
  }
  await page.waitForSelector('#search-box', { timeout: 10000 });
  await page.evaluate(() => { document.querySelector('#search-box').value = ''; });
  await page.type('#search-box', courseCode);
  console.log(`🔍 Searched for course: ${courseCode}`);
  await sleep(1000);
  const firstResultXPath = "//*[@id='0_']/a";
  await clickXPath(page, firstResultXPath, "first course result");
  return true;
}

async function main() {
  try {
    const selectedFile = await askForFile();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(selectedFile);
    const sheet = workbook.worksheets[0];

    const browser = await puppeteer.launch({ headless: false, defaultViewport: null, args: ['--start-maximized'] });
    const page = await browser.newPage();
    await page.goto('https://york-sbx.kuali.co/cor/main/#/apps');
    console.log("🟢 Logged in? Please log in manually if prompted.");
    await page.waitForFunction(() => !location.href.includes('passportyork'), { timeout: 0 });
    console.log("✅ Login complete. Waiting for page to load...");
    await sleep(15000);

    await navigateToCourses(page);

    let lastCourseCode = null;
    let currentCourseRows = [];

    // First, group all rows by course code
    const courseGroups = {};
    for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex++) {
      const row = sheet.getRow(rowIndex);
      const faculty = row.getCell('A').value;
      const dept = row.getCell('C').value;
      const code = row.getCell('D').value;
      
      if (!faculty || !dept || !code) {
        console.log(`⚠️ Skipping row ${rowIndex} due to missing course information.`);
        continue;
      }
      
      const courseCode = `${faculty}/${dept} ${code}`.replace(/\s+/g, ' ').trim();
      if (!courseGroups[courseCode]) {
        courseGroups[courseCode] = [];
      }
      courseGroups[courseCode].push({
        rowIndex,
        clo: row.getCell('G').value,
        gai: String(row.getCell('I').text).trim(),
        gaml: row.getCell('L').value
      });
    }

    // Now process each course
    for (const courseCode in courseGroups) {
      console.log(`🔄 Switching to course: ${courseCode}`);
      await clickXPath(page, "//*[@id='app']/div/div[4]/nav/ul/li[3]/a/img", "Courses");
      await sleep(1000);
      const success = await searchCourse(page, courseCode);
      if (!success) continue;
      await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[1]/div/div[1]/a", "Edit");
      await clickXPath(page, "//*[@id='app']/div/div[4]/div/main/div/div[3]/div[1]/div/div[3]/nav/ul/li[10]/div", "Lassonde Course Outcomes");
      await sleep(1000);

      // Delete all existing CLOs before adding new ones
      await deleteAllExistingCLOs(page);
      
      // Add all CLOs for this course
      for (const rowData of courseGroups[courseCode]) {
        await inputNewCLO(page, rowData.clo, rowData.gai, rowData.gaml);
      }
    

      // Delete the last CLO for this course
      const deleteButtonXPath = "//button[.//i[contains(@class, 'fa-trash')]]";
      const deleteButtons = await page.$x(deleteButtonXPath);
      if (deleteButtons.length > 0) {
        const lastDeleteButton = deleteButtons[deleteButtons.length - 1];
        await lastDeleteButton.click();
        console.log(`✅ Deleted the last CLO for ${courseCode}.`);
        await sleep(1000); // Wait for deletion to complete
      } else {
        console.log(`⚠️ No delete buttons found for ${courseCode}.`);
      }
    }

    console.log("✅ All courses processed. Ready for review.");
  } catch (err) {
    console.error("❌ Error in automation:", err);
  }
}

main();
