/**
 * Calculates the variability score of a given array of text.
 *
 * This function uses a simple approach:
 * 1. Calculate the average word length of all texts.
 * 2. Calculate the standard deviation of word lengths.
 * 3. Calculate the average number of unique words in each text.
 * 4. Calculate the standard deviation of unique word counts.
 * 5. Combine the above metrics into a single score.
 * 6. Normalize the score based on the mean and standard deviation of a reference distribution.
 * 7. Calculate the final score based on the distribution of the reference bell curve.
 *
 * @param {string[]} textArray - An array of text strings.
 * @returns {number} - The final variability score.
 */
function calculateMappedVariabilityScore(textArray) {
  if (!textArray || textArray.length === 0) {
    return 0; // Handle empty arrays
  }

  const wordLengths = textArray.map(text => {
    const words = text.split(/\s+/);
    return words.map(word => word.length).reduce((sum, length) => sum + length, 0) / words.length;
  });

  const avgWordLength = wordLengths.reduce((sum, length) => sum + length, 0) / wordLengths.length;
  const stdDevWordLength = calculateStandardDeviation(wordLengths);

  const uniqueWordCounts = textArray.map(text => {
    const words = text.toLowerCase().split(/\s+/);
    return new Set(words).size;
  });

  const avgUniqueWords = uniqueWordCounts.reduce((sum, count) => sum + count, 0) / uniqueWordCounts.length;
  const stdDevUniqueWords = calculateStandardDeviation(uniqueWordCounts);

  // Combine metrics into a single score (adjust weights as needed)
  const variabilityScore = 
    (stdDevWordLength * 0.5) + 
    (stdDevUniqueWords * 0.5); 

  // Mean and standard deviation of the old and new scores
  const mean_old = 15.732;
  const std_dev_old = 6.545;

  const mean_new = 38.8;
  const std_dev_new = 9.8;

  // Normalize the variability score based on the mean and standard deviation of the old scores
  const normalizedVariabilityScore = (variabilityScore - mean_old) / std_dev_old;

  // Calculate the final score based on the distribution of the new scores
  const finalScore = mean_new + (normalizedVariabilityScore * std_dev_new);

  return finalScore;
}

function calculateStandardDeviation(values) {
  if (!Array.isArray(values) || values.length === 0) {
    console.warn('Invalid input for standard deviation calculation.');
    return 0;
  }

  const avg = values.reduce((sum, value) => sum + value, 0) / values.length;
  const squaredDifferences = values.map(value => Math.pow(value - avg, 2));
  const variance = squaredDifferences.reduce((sum, diff) => sum + diff, 0) / values.length;
  return Math.sqrt(variance);
}

function getTableData(table) {
  if (!table) {
    return []; // Handle null or undefined tables
  }

  const rows = table.getNumRows();
  const cols = table.getRow(0).getNumCells(); // Use the first row to determine the number of columns
  const data = [];

  for (let i = 0; i < rows; i++) {
    const row = [];
    for (let j = 0; j < cols; j++) {
      const cell = table.getRow(i).getCell(j); // Updated to use getRow and getCell
      row.push(cell.getText());
    }
    data.push(row);
  }

  return data;
}

function getDocTables(fileIds) {
  const docTables = [];

  for (const fileId of fileIds) {
    try {
      const doc = DocumentApp.openById(fileId); 
      const body = doc.getBody();
      const tables = body.getTables();

      const docData = { 
        fileId: fileId, 
        tables: [] 
      };

      if (tables.length > 0) { 
        for (const table of tables) {
          docData.tables.push(getTableData(table)); 
        }
      } else {
        console.warn(`No tables found in document with ID: ${fileId}`);
      }

      docTables.push(docData);
    } catch (error) {
      // Handle potential errors (e.g., document not found, permission issues)
      console.error(`Error accessing document with ID: ${fileId}`, error);
    }
  }

  return docTables;
}

function analyzeTables() {
  const fileIds = ["1kFNwcPSso9lyCBG2wn2uVSKCtTLye4SH_CsqTPtayI4"];

  for (const fileId of fileIds) {
    const docTablesData = getDocTables([fileId]);

    if (docTablesData.length > 0 && docTablesData[0].tables.length > 0) {
      const doc = DocumentApp.openById(fileId);
      const body = doc.getBody();
      const tables = body.getTables();

      // Iterate over all tables in the document
      tables.forEach((table, index) => {
        if (table) {
          const lastRowIndex = table.getNumRows() - 1;
          if (lastRowIndex >= 0) {
            const lastRow = table.getRow(lastRowIndex);
            const cellTexts = [];

            // Collect text from the last row
            for (let colIndex = 0; colIndex < lastRow.getNumCells(); colIndex++) {
              cellTexts.push(lastRow.getCell(colIndex).getText());
            }

            // Calculate mapped variability score
            const variabilityScore = calculateMappedVariabilityScore(cellTexts);

            // Add a comment to the first cell in the last row
            const firstCell = lastRow.getCell(0);

            // Create a range for the first cell
            const rangeBuilder = doc.newRange();
            rangeBuilder.addElement(firstCell);
            const range = rangeBuilder.build();

            try {
              range.getRangeElements()[0].getElement().asText().editAsText()
                .setText(`Mapped Variability Score: ${variabilityScore.toFixed(2)}`);
              console.log(`Comment added to table ${index + 1}: Mapped Variability Score: ${variabilityScore.toFixed(2)}`);
            } catch (error) {
              console.error(`Error adding comment to table ${index + 1} in document:`, error);
            }
          }
        }
      });
    } else {
      console.log(`No tables found in document with ID: ${fileId}`);
    }
  }
}
