const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const { OpenAI } = require('openai');
const mysql = require('mysql2/promise');
require('dotenv').config();
const openai = new OpenAI({
  apiKey: process.env.TCSION_AI_KEY,
});

const ExcelJS = require('exceljs');

async function parseExcelFile(filePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];

    // Get header row (first row)
    const headerRow = worksheet.getRow(1);
    const headers = headerRow.values.slice(0); // remove empty first element

    const expectedHeaders = [
      'IssueId',
      'Department',
      'DataExistsIn',
      'usersAffected',
      'ImpactedData',
      'issueSummary',
      'issueStatus',
      'IssuePriority',
      'IssueDescription',
      'StepsToReproduceDefects',
      'URL',
      'DefectsFoundDuringTheseScenarios',
      'RegressionRequired'
    ];

    // Validate headers
    if (headers.length !== 14) {
      throw new Error('Invalid file: Expected 13 columns');
    }

    const data = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowData = {};

      expectedHeaders.forEach((header, index) => {
        const cell = row.getCell(index + 1);
        const value = cell!==null && cell.text ? cell.text : '';
        rowData[header] = value ? String(value).trim() : '';
      });

      // Skip only if entire row is empty
      //const isEmptyRow = Object.values(rowData).every(val => val === '');

      if(rowNumber==413){
      console.log(rowNumber);
      }
      //if (!isEmptyRow) {
        data.push(rowData);
      //}
    });

    return data;

  } catch (error) {
    console.error('Error parsing Excel:', error);
    throw error;
  }
}

// Usage within your async workflow


(async () => {
  try {
    const items = await parseExcelFile('service-now.xlsx');
    //const allRows = items.map(item => item?.json ?? item);

    //const rowCount = allRows.length;

   // let columns = Object.keys(allRows[0] || {});
   // columns = columns.map(c => c.replace(/^\uFEFF/, ""));
    //let missingStats = {};

    //let typeStats = {};

    //let basicStats = {};

    // columns.forEach(col => {

    //     let values = allRows.map(r => r[col]);
    //     let missing = values.filter(v => v === null || v === "" || v === undefined).length;

    //     missingStats[col] = ((missing / rowCount) * 100).toFixed(2);
    //     let numericCount = values.filter(v => !isNaN(parseFloat(v))).length;

    //     typeStats[col] = {
    //       numeric_percentage: ((numericCount / rowCount) * 100).toFixed(2)
    //     };

    //     let negatives = values.filter(v => !isNaN(v) && Number(v) < 0).length;
    //     basicStats[col] = {
    //       negative_values: negatives
    //     };
    // });
    
    //const datasample = [{json: {row_count: rowCount,columns: columns,missing_percentage: missingStats,
    //  type_analysis: typeStats,anomaly_signals: basicStats,sample_rows: allRows.slice(0, 13)}}];

    //let rawText = await runDataQualityAgent(datasample);

    //if (!rawText) {
    //   rawText = [{ json: { error: "OpenAI text not found", received: $json } }];
    //}

    //let aiOutput;
    //  try {
        //aiOutput = JSON.parse(rawText.choices[0].message.content);
    //  } catch (err) {
    //    console.error(err);
    //  }

      // if (!aiOutput.missing_values) {
      //   aiOutput = [{ json: { error: "missing_values_analysis not found", aiOutput } }];
      // }

      //const missingValues = aiOutput.missing_values;
      //let maxMissing = 0;
      //for (const col in missingValues) {
      //  const percent = parseFloat(String(missingValues[col]).replace("%", ""));
      //  if (!isNaN(percent) && percent > maxMissing) {
      //    maxMissing = percent;
      // }
      //}
      //rawText = [{json: {max_missing_percentage: maxMissing,missing_values: missingValues}}];
      
      //let value = maxMissing;

      //if (value > 20) {
      //  return [{
      //    json: {
      //      workflow_status: "Stopped",
      //      reason: "Missing values exceed acceptable threshold"
      //    }
      //  }];
      //}

      const cleaned = [];

      for (const item of items) {
        const row = item;
        if (row.IssueId === null || row.IssueId === undefined || String(row.IssueId).trim() === "") {
          continue;
        }
        
        // Trim whitespace from string fields
        row.Department = String(row.Department).trim() || null;
        row.DataExistsIn = String(row.DataExistsIn).trim() || null;
        row.usersAffected = String(row.usersAffected).trim() || null;
        row.ImpactedData = String(row.ImpactedData).trim() || null;
        row.IssueId = String(row.IssueId).trim() || null;
        row.issueSummary = String(row.issueSummary).trim() || null;
        row.IssuePriority = String(row.IssuePriority).trim() || null;
        row.IssueDescription = String(row.IssueDescription).trim() || null;
        row.StepsToReproduceDefects = String(row.StepsToReproduceDefects).trim() || null;
        row.DefectsFoundDuringTheseScenarios = String(row.DefectsFoundDuringTheseScenarios).trim() || null;
        row.URL = String(row.URL).trim() || null;
        row.RegressionRequired = String(row.RegressionRequired).trim() || null;
        //row.issueSummary = String(row.issueSummary).trim();
        cleaned.push(row);    
    } 
        await upsertServiceNow(cleaned); 
              
        const data = {
          'overview': await getDefectStatsForRegressionRequiredByDepartmentOnTopPriority(),
          'defectsByDepartmentAndUsersAffected': await getDefectCountsByDepartmentAndUsersAffected()
        };
        await generateExecutiveReport(data);

} catch (err) {
    console.error(err);
  }
})();

/**
 * Node Name: Data Quality Agent
 * Model: gpt-4.1-mini
 */
async function runDataQualityAgent(datasample) {

  console.log(`Running Data Quality Agent with sample data: ${JSON.stringify(datasample)}`);
  const response = await openai.chat.completions.create({
  model: "gpt-4.1-mini",
  messages: [
    {
      role: "system",
      content: `You are a Data Quality Analysis Agent.

Rules:
1. Only refer to provided column names.
2. Do not invent new columns.
3. If insufficient evidence exists, state: "Not detectable from provided data".
Return ONLY JSON with this exact structure:
    {
      "missing_values": {}
    }

    Do NOT rename these keys.
    Do NOT add new keys.
    `
    },
    {
      role: "user",
      content: '${JSON.stringify(datasample)}'
    }
  ],
  response_format: { type: "json_object" }
});

  const result = JSON.parse(response.choices[0].message.content);
  console.log(`Data Quality Analysis Result: ${JSON.stringify(result)}`);
  return response;
}



async function upsertServiceNow(cleaned) {
  const connection = await mysql.createConnection({
    host: process.env.DB_HOST,
    user: 'face',
    password: process.env.DB_PASS,
    database: process.env.DB_NAME
  });

  // The query: note the use of VALUES() in the update section
  // to refer to the data we attempted to insert.
  const sql = `
        INSERT INTO servicenow (IssueId,Department,DataExistsIn,usersAffected,ImpactedData,issueSummary,issueStatus,IssuePriority,IssueDescription,StepsToReproduceDefects,URL,DefectsFoundDuringTheseScenarios,RegressionRequired)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
            IssueId = VALUES(IssueId),
            Department = VALUES(Department),
            DataExistsIn = VALUES(DataExistsIn),
            usersAffected = VALUES(usersAffected),
            ImpactedData = VALUES(ImpactedData),
            issueSummary = VALUES(issueSummary),
            issueStatus = VALUES(issueStatus),
            IssuePriority = VALUES(IssuePriority),
            IssueDescription = VALUES(IssueDescription),
            StepsToReproduceDefects = VALUES(StepsToReproduceDefects),
            URL = VALUES(URL),
            DefectsFoundDuringTheseScenarios = VALUES(DefectsFoundDuringTheseScenarios),
            RegressionRequired = VALUES(RegressionRequired)
    `;

  try {
    for (const row of cleaned) {
        await connection.execute(sql, [
          row.IssueId,
          row.Department,
          row.DataExistsIn,
          row.usersAffected,
          row.ImpactedData,
          row.issueSummary,
          row.issueStatus,
          row.IssuePriority,
          row.IssueDescription,
          row.StepsToReproduceDefects,
          row.URL,
          row.DefectsFoundDuringTheseScenarios,
          row.RegressionRequired
        ]);
      }
  } catch (err) {
    console.error('Error executing upsert:', err);
  }
}

async function getDefectStatsForRegressionRequiredByDepartmentOnTopPriority() {
  // 1. Establish connection
  const connection = await mysql.createConnection({
    host: process.env.DB_HOST,
    user: 'face',
    password: process.env.DB_PASS,
    database: process.env.DB_NAME
  });

  const sql = `
        SELECT 
    Department,
    COUNT(*) AS defect_count,
    SUM(CASE WHEN IssuePriority IN ('High','Highest') THEN 1 ELSE 0 END) AS high_priority_count,
    SUM(CASE WHEN IssuePriority IN ('Lowest','Low','Medium') THEN 1 ELSE 0 END) AS low_priority_count,
    COUNT(DISTINCT issueSummary) AS unique_issues,
    COUNT(*) - COUNT(DISTINCT issueSummary) AS repeated_issues
FROM servicenow
GROUP BY Department
HAVING defect_count > 1
ORDER BY high_priority_count DESC, repeated_issues DESC, defect_count DESC;
    `;

  try {
    // 2. Execute query
    const [rows] = await connection.query(sql);

    // 3. Convert the first row to a JSON string (if needed) 
    // or just work with the object directly
    const stats = rows[0];

    console.log('--- Raw Object ---');
    console.log(stats);

    console.log('\n--- Formatted JSON ---');
    console.log(JSON.stringify(stats, null, 2));

    return stats;

  } catch (err) {
    console.error('Error fetching stats:', err);
  } finally {
    await connection.end();
  }
}

async function getDefectCountsByDepartmentAndUsersAffected() {
  const connection = await mysql.createConnection({
    host: process.env.DB_HOST,
    user: 'face',
    password: process.env.DB_PASS,
    database: process.env.DB_NAME
  });

  const sql = `
        SELECT 
    Department,
    usersAffected,
    COUNT(*) AS defect_count
FROM servicenow
GROUP BY Department, usersAffected
ORDER BY defect_count DESC;
    `;

  try {
    const [rows] = await connection.query(sql);

    // 'rows' will be an array of objects
    //console.log("Department Statistics:");
    //console.table(rows); // console.table is great for viewing grouped data!

    return rows;
  } catch (err) {
    console.error("Query failed:", err);
  } 
}

async function generateExecutiveReport(data) {
  const overviewStr = JSON.stringify(data.overview);
  const deptStr = JSON.stringify(data.defectsByDepartmentAndUsersAffected);

  const prompt = `You are a senior QA analytics expert. Generate a concise executive-level defect report based on the provided JSON data.

Context:
- "overview" contains defect metrics grouped by Department:
  - defect_count
  - high_priority_count (High + Highest)
  - low_priority_count (Lowest + Low + Medium)
  - unique_issues
  - repeated_issues

- "defectsByDepartmentAndUsersAffected" shows defect distribution based on user impact.

Your Goals:
1. Identify high-risk departments for regression testing.
2. Highlight departments with:
   - high defect volume
   - high priority defects
   - repeated issues (stability concerns)
3. Analyze user impact and identify business-critical areas.
4. Provide clear, actionable insights for leadership.

Output Format:

1. Executive Summary (5–6 lines, non-technical)
2. Key Insights
   - Top 3 high-risk departments (with reasons)
   - Stability concerns (repeated issues)
   - Severity distribution insights
3. User Impact Analysis
   - Which departments affect the most users
   - Business risk explanation
4. Regression Testing Recommendations
   - Priority areas
   - Suggested focus scenarios
5. Optional: Risk Score (Lowest / Low / Medium / High per department)

Guidelines:
- Keep it concise and professional
- Avoid raw SQL or technical jargon
- Use bullet points where helpful
- Base conclusions strictly on the data provided (no assumptions)

Here is the data:
{{data}}`.replace('{{data}}', JSON.stringify(data));;

  try {
    const json = await openai.chat.completions.create({
      model: "gpt-4.1-mini",
      messages: [
        { role: "user", content: prompt }
      ]
    });

    const result = json.choices?.[0]?.message?.content;
    console.log(result);
    //console.log(`Data Quality Analysis Result: `);
    //console.log('Generated Report:\n ${JSON.stringify(result, null, 2)}');
    const fsp = require('fs/promises');
    await fsp.writeFile("Service-Now.txt","");
    await fsp.writeFile(
      `Service-Now.txt`,
       result
    );

  } catch (error) {
    console.error("Error calling AI model:", error);
  }
}