async function extractData() {
  const fileUrl =
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQE4EsYXPu7gSCeMt0zO71d-zisc2wPQ6Xem7R5oTdo6lTHxdSa9KZu6DbOHYyGVA/pub?output=csv"; // Replace with your CSV link

  try {
    // Fetch CSV data
    const response = await fetch(fileUrl);
    const text = await response.text();

    // Convert CSV to JSON
    const rows = text.split("\n").map((row) => row.split(","));
    const headers = rows[0]; // First row as headers
    const data = rows
      .slice(1)
      .map((row) =>
        Object.fromEntries(
          row.map((value, index) => [headers[index].trim(), value.trim()])
        )
      );

    // Read the USN input
    let usn = document.getElementById("USNInput").value.trim().toUpperCase();

    // Check if USN is empty
    if (!usn) {
      alert("Please enter a valid USN.");
      return;
    }

    // Search for the USN in the data
    let found = false;
    let record = null;

    for (let i = 0; i < data.length; i++) {
      if (usn === data[i].USN.trim().toUpperCase()) {
        found = true;
        record = data[i]; // Store the matched record
        break; // Exit loop once a match is found
      }
    }

    // Remove old result or error message if any
    const oldMessage = document.getElementById("result-message");
    if (oldMessage) oldMessage.remove();

    const resultContainer = document.createElement("div");
    resultContainer.setAttribute("id", "result-message");
    resultContainer.style.marginTop = "20px"; // Ensure there’s spacing from top of the page

    if (!found) {
      // No match found
      resultContainer.innerHTML = "<h3>NO RECORD FOUND</h3>";
      resultContainer.style.padding = "5px";
      resultContainer.style.textAlign = "center";
      resultContainer.style.marginTop = "10px";
    } else {
      // Display Name and USN above the table
      const nameUSNContainer = document.createElement("div");
      nameUSNContainer.innerHTML = `
        <p><strong>Name:</strong> ${record.Name}</p>
        <p><strong>USN:</strong> ${record.USN}</p>
      `;
      nameUSNContainer.style.marginBottom = "20px"; // More space between name/USN and table
      nameUSNContainer.style.textAlign = "center"; // Center the name and USN
      nameUSNContainer.style.display = "block"; // Ensure it's a block-level element
      resultContainer.appendChild(nameUSNContainer); // Append the Name and USN above the table

      // Create the table for marks and SGPA
      let r_table = document.createElement("table");
      r_table.style.width = "70%"; // Limit the table width
      r_table.style.margin = "0 auto"; // Center the table horizontally
      r_table.style.marginTop = "15px"; // Space between name and table
      r_table.style.borderCollapse = "collapse"; // For better table appearance
      r_table.style.textAlign = "center"; // Align table content to center
      r_table.style.border = "1px solid black"; // Border for the table

      // Table header
      r_table.insertAdjacentHTML(
        "beforeend",
        "<tr><th>SI.No</th><th>Subject</th><th>Marks</th></tr>"
      );

      // Table rows
      for (let j = 1, k = 7; j <= 8; j++, k++) {
        let new_row = document.createElement("tr");
        new_row.insertAdjacentHTML(
          "beforeend",
          `<td>${j}</td><td>Subject ${j}</td><td>${
            record[Object.keys(record)[k]]
          }</td>`
        );
        r_table.appendChild(new_row); // Append the new row to the table
      }

      // SGPA row
      let new_row = document.createElement("tr");
      new_row.insertAdjacentHTML(
        "beforeend",
        `<td></td><td>SGPA</td><td>${record.SGPA}</td>`
      );
      r_table.appendChild(new_row);

      // Append the table to the result container
      resultContainer.appendChild(r_table);

      // Create Print Result Button
      const printButton = document.createElement("button");
      printButton.innerText = "Print Result";
      printButton.style.backgroundColor = "#4CAF50"; // Green background color
      printButton.style.color = "white"; // White text
      printButton.style.padding = "10px 20px"; // Button padding
      printButton.style.border = "none"; // Remove border
      printButton.style.borderRadius = "5px"; // Rounded corners
      printButton.style.marginTop = "20px"; // Space above button
      printButton.style.cursor = "pointer"; // Pointer cursor on hover
      printButton.style.display = "block"; // Block-level button
      printButton.style.marginLeft = "auto"; // Center the button horizontally
      printButton.style.marginRight = "auto";

      // Add event listener to print
      printButton.addEventListener("click", function () {
        const printWindow = window.open("", "_blank"); // Open a new tab
        printWindow.document.write("<html><head><title>Print Result</title>");

        // Add print-specific styles
        printWindow.document.write(`
          <style>
            body {
              font-family: Arial, sans-serif;
            }
            table {
              width: 100%;
              border-collapse: collapse;
            }
            table, th, td {
              border: 1px solid black;
            }
            th, td {
              padding: 8px;
              text-align: center;
            }
            h1, h3 {
              text-align: center;
            }
          </style>
        `);

        printWindow.document.write("</head><body>");
        printWindow.document.write(resultContainer.innerHTML); // Write content to new tab
        printWindow.document.write("</body></html>");
        printWindow.document.close();
        printWindow.print(); // Trigger print dialog
      });

      // Append the Print Result button
      resultContainer.appendChild(printButton);
    }

    // Append the message to the main container
    document.getElementById("dt").after(resultContainer);
  } catch (error) {
    console.error("Error loading the file:", error);
    alert("Failed to load the Excel file.");
  }
}

// Attach event listener to the button
document.getElementById("submit-btn").addEventListener("click", extractData);
