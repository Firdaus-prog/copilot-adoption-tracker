import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import csvParser from 'csv-parser';

const inputFolder = path.resolve(__dirname, '../msFiles');
const outputFolder = path.resolve(__dirname, '../result');
const excelOutputPath = path.join(outputFolder, 'department_analysis.xlsx');

const filePattern = /^copilot_users_(\d{8}).csv$/;

const allCategories = [
  'Content',
  'Customer Service',
  'Human Resources',
  'Finance',
  'payTV',
  'Marketing',
  'Communication & Sustainability',
  'Astro Audio',
  'Uncategorized',
];

// Helper: Convert "20250801" to "1st August"
function formatDateLabel(dateStr: string): string {
  const year = parseInt(dateStr.slice(0, 4), 10);
  const month = parseInt(dateStr.slice(4, 6), 10) - 1;
  const day = parseInt(dateStr.slice(6, 8), 10);

  const suffix =
    day === 1 ? 'st' :
    day === 2 ? 'nd' :
    day === 3 ? 'rd' : 'th';

  const monthName = new Date(year, month, day).toLocaleString('en-GB', { month: 'long' });
  return `${day}${suffix} ${monthName}`;
}

// Collectors
const categorized: Record<string, Set<string>> = {};
const allDateCounts: Record<string, Record<string, number>> = {};

allCategories.forEach(cat => {
  categorized[cat] = new Set();
});

// Process one CSV file
function processCSV(filePath: string, date: string): Promise<void> {
  const fileCategoryCount: Record<string, number> = {};
  allCategories.forEach(cat => fileCategoryCount[cat] = 0);

  return new Promise((resolve, reject) => {
    fs.createReadStream(filePath)
      .pipe(csvParser())
      .on('data', (row) => {
        const activity1 = row["Last activity date of Copilot.cloud.microsoft (UTC)"]?.trim();
        const activity2 = row["Last activity date of Microsoft 365 Copilot (app) (UTC)"]?.trim();
        const department = row["department"]?.toString().trim();

        if ((activity1 || activity2) && department) {
          const category = mapToCategory(department);
          categorized[category].add(department);
          fileCategoryCount[category]++;
        }
      })
      .on('end', () => {
        allDateCounts[date] = fileCategoryCount;
        resolve();
      })
      .on('error', reject);
  });
}

// Categorize departments
function mapToCategory(dept: string): string {
  if (!dept) return 'Uncategorized';
  const lower = dept.toLowerCase();

  if (/content|programming|editorial|broadcast|production|creative|media|post|studio|shaw|vod|tutor tv|magazine|video|visual|design|copywriting|rojak|thinker/i.test(lower)) return 'Content';
  if (/customer|call centre|ccc|service recovery|sales support|customer experience|relationship/i.test(lower)) return 'Customer Service';
  if (/hr|employee|payroll|talent|learning|industrial relations|legal&hr|legal division|employee engagement/i.test(lower)) return 'Human Resources';
  if (/finance|cfo|ap & ar|tax|reporting|treasury|corporate finance|account|admin/i.test(lower)) return 'Finance';
  if (/paytv|pay tv|astro ria|astro prima|astro warna|njoy|sooka/i.test(lower)) return 'payTV';
  if (/marketing|promo|digital marketing|base marketing|product marketing|social media|retention|winback|liaison|trade/i.test(lower)) return 'Marketing';
  if (/communication|regulatory|corporate affairs|stakeholder|govt|strategy|sustainability|esg|public/i.test(lower)) return 'Communication & Sustainability';
  if (/audio|radio|gegar|ceria|amp|astro audio/i.test(lower)) return 'Astro Audio';

  return 'Uncategorized';
}

// Main function
async function runAnalysis() {
  try {
    if (!fs.existsSync(outputFolder)) {
      fs.mkdirSync(outputFolder);
    }

    const files = fs.readdirSync(inputFolder).filter(file => filePattern.test(file));
    if (files.length === 0) {
      console.log('❌ No matching CSV files found.');
      return;
    }

    for (const file of files) {
      const match = file.match(filePattern);
      if (!match) continue;

      const date = match[1]; // e.g. 20250801
      const fullPath = path.join(inputFolder, file);
      console.log(`📄 Processing ${file}`);
      await processCSV(fullPath, date);
    }

    const sortedDates = Object.keys(allDateCounts).sort();
    const formattedHeaders = sortedDates.map(formatDateLabel);

    // Sheet: Category Counts
    const categoryCountsSheet: any[][] = [["Category", ...formattedHeaders]];
    for (const category of allCategories) {
      const row = [category];
      for (const date of sortedDates) {
        row.push(allDateCounts[date][category] || 0);
      }
      categoryCountsSheet.push(row);
    }

    // Sheet: Department List
    const deptSheet: any[][] = [["Category", "Department"]];
    for (const category of allCategories) {
      const depts = Array.from(categorized[category]).sort();
      for (const dept of depts) {
        deptSheet.push([category, dept]);
      }
    }

    // Write Excel
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(categoryCountsSheet), "Category Counts");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(deptSheet), "Department List");

    XLSX.writeFile(wb, excelOutputPath);
    console.log(`✅ Excel file written to ${excelOutputPath}`);
  } catch (err) {
    console.error('❌ Error:', err);
  }
}

runAnalysis();
