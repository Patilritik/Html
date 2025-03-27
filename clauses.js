console.log("Before Office.onReady");

Office.onReady(async () => {
    console.log("Office.onReady triggered");

    if (!localStorage.getItem("loginData")) {
        console.log("No login data found");
        window.location.href = "index.html";
        return;
    }

    const loginData = JSON.parse(localStorage.getItem("loginData"));
    const UserId = loginData.loginid;
    const ComCode = loginData.comcode;
    const status = loginData.status;
    const Token = loginData.Token;

    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");
    const proceedBtn = document.getElementById("proceedBtn");
    const loader = document.getElementById("loader");

    function showLoader() {
        loader.style.display = "block";
    }

    function hideLoader() {
        loader.style.display = "none";
    }

    async function fetchDepartmentList() {
        const apiUrl = `https://lapi.convergelego.com/api/AddLegalAgreement/Departmentlisit?companycode=${ComCode}&status=${status}`;
        try {
            const response = await fetch(apiUrl, {
                method: "GET",
                headers: {
                    "Userid": UserId,
                    "Key": "0",
                    "Comcode": ComCode
                }
            });

            const result = await response.json();
            departmentSelect.innerHTML = "";

            if (result.status === 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Department";
                departmentSelect.appendChild(defaultOption);

                const departments = result?.Detail?.data || [];
                departments.forEach(dept => {
                    const option = document.createElement("option");
                    option.value = dept.DeptId;
                    option.textContent = dept.DeptName;
                    departmentSelect.appendChild(option);
                });

                if (loginData.Depid) {
                    departmentSelect.value = loginData.Depid;
                    fetchAgreementTypeList(loginData.Depid);
                }
            } else {
                departmentSelect.innerHTML = `<option value="">No departments found</option>`;
            }
        } catch (err) {
            console.error("Department fetch error:", err);
            departmentSelect.innerHTML = `<option value="">No departments found</option>`;
        }
    }

    async function fetchAgreementTypeList(deptId) {
        const apiUrl = `https://lapi.convergelego.com/api/STDAgreementType/MstAgTypelisit?comcode=${ComCode}&depid=${deptId}`;
        try {
            agreementTypeSelect.disabled = false;
            const response = await fetch(apiUrl, {
                method: "GET",
                headers: {
                    "Userid": UserId,
                    "Key": "0",
                    "Comcode": ComCode
                }
            });

            const result = await response.json();
            agreementTypeSelect.innerHTML = "";

            if (result.status === 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Agreement Type";
                agreementTypeSelect.appendChild(defaultOption);

                const types = result?.Detail?.data || [];
                types.forEach(ag => {
                    const option = document.createElement("option");
                    option.value = ag.Agtypeid;
                    option.textContent = ag.AgtypeDesc;
                    agreementTypeSelect.appendChild(option);
                });
            } else {
                agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
            }
        } catch (err) {
            console.error("Agreement type fetch error:", err);
            agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
        }

        updateProceedButtonState();
    }

    function updateProceedButtonState() {
        proceedBtn.disabled = !(departmentSelect.value && agreementTypeSelect.value);
    }
    
    // Table format
    // async function copyToWord(clauses) {
    //     try {
    //         await Word.run(async (context) => {
    //             console.log("Copying to Word:", clauses);
    //             const body = context.document.body;
                
    //             // Insert table with rows and columns
    //             const table = body.insertTable(clauses.length + 1, 5, Word.InsertLocation.end);
                
    //             // Load table properties
    //             table.load("values, style, columns");

    //             await context.sync();

    //             // Set header row
    //             table.values[0] = ["Clause ID", "Title", "Description", "Created By", "Created On"];

    //             // Set data rows
    //             clauses.forEach((clause, index) => {
    //                 table.values[index + 1] = [
    //                     clause.id || '-',
    //                     clause.causetitle || '-',
    //                     clause.cause || '-',
    //                     clause.crby || '-',
    //                     clause.cron || '-'
    //                 ];
    //             });
    //             console.log("table",table);
    //             console.log("table.values",table.values);
    //             // Apply formatting
    //             table.style = "Grid Table 4 - Accent 1";
    //             table.getRange().font.size = 10;

    //             // Set column widths (in points)
    //             const columns = table.columns;
    //             console.log("columns",columns);
    //             console.log("columns.items",columns.items);
    //             console.log("columns.items[0]",columns.items[0]);
    //             columns.items[0].setWidth(60, Word.WidthUnits.points);  // Clause ID
    //             columns.items[1].setWidth(100, Word.WidthUnits.points); // Title
    //             columns.items[2].setWidth(200, Word.WidthUnits.points); // Description
    //             columns.items[3].setWidth(80, Word.WidthUnits.points);  // Created By
    //             columns.items[4].setWidth(80, Word.WidthUnits.points);  // Created On

    //             await context.sync();
    //         });
    //     } catch (error) {
    //         console.error("Error copying to Word:", error);
    //         alert("Error copying to Word document: " + error.message);
    //     }
    // }

    // Checking Table format
    // async function copyToWord(clauses) {
    //     try {
    //         await Word.run(async (context) => {
    //             console.log("Copying to Word:", clauses);
    //             const body = context.document.body;
                
    //             // Insert table with rows and columns
    //             const table = body.insertTable(clauses.length + 1, 5, Word.InsertLocation.end);
                
    //             // Load table properties
    //             table.load("values, style, columns");

    //             await context.sync();

    //             // Set header row
    //             table.values[0] = ["Clause ID", "Title", "Description", "Created By", "Created On"];

    //             // Set data rows
    //             clauses.forEach((clause, index) => {
    //                 table.values[index + 1] = [
    //                     clause.id || '-',
    //                     clause.causetitle || '-',
    //                     clause.cause || '-',
    //                     clause.crby || '-',
    //                     clause.cron || '-'
    //                 ];
    //             });
    //             console.log("table",table);
    //             console.log("table.values",table.values);
    //             // Apply formatting
    //             table.style = "Grid Table 4 - Accent 1";
    //             table.getRange().font.size = 10;

    //             // Try to set column widths
    //         try {
    //             const columns = table.columns;
    //             table.load("columns"); // Explicitly load columns
    //             await context.sync(); // Sync again to ensure columns are available
    //             console.log("Columns loaded:", columns);
                
    //             if (columns && columns.items && columns.items.length > 0) {
    //                 columns.items[0].setWidth(60, Word.WidthUnits.points);  // Clause ID
    //                 columns.items[1].setWidth(100, Word.WidthUnits.points); // Title
    //                 columns.items[2].setWidth(200, Word.WidthUnits.points); // Description
    //                 columns.items[3].setWidth(80, Word.WidthUnits.points);  // Created By
    //                 columns.items[4].setWidth(80, Word.WidthUnits.points);  // Created On
    //             } else {
    //                 console.warn("Columns not available, skipping width adjustment");
    //             }
    //         } catch (colError) {
    //             console.warn("Column width adjustment failed:", colError);
    //             // Fallback: Use table preferred width and auto-fit
    //             table.preferredWidth = 520; // Total width in points (60 + 100 + 200 + 80 + 80)
    //             table.autoFitContent();
    //         }

    //         await context.sync();
    //         console.log("Table successfully created");
    //         });
    //     } catch (error) {
    //         console.error("Error copying to Word:", error);
    //         alert("Error copying to Word document: " + error.message);
    //     }
    // }

    // Workable
    // async function copyToWord(clauses) {
    //     try {
    //       await Word.run(async (context) => {
    //         const body = context.document.body;
    //         const rowCount = clauses.length + 1; // +1 for header row
    //         const colCount = 5;
      
    //         // Insert table at the end of the document
    //         const table = body.insertTable(rowCount, colCount, Word.InsertLocation.end);
    //         table.style = "Grid Table 4 - Accent 1";
    //         table.getRange().font.size = 12;
      
    //         // Debug: log clauses to verify content
    //         console.log("Clauses to insert:", clauses);
      
    //         // Prepare table data (headers + rows)
    //         const values = [
    //           ["Clause ID", "Title", "Description", "Created By", "Created On"],
    //           ...clauses.map(clause => [
    //             clause?.id || "-",
    //             clause?.causetitle || "-",
    //             clause?.cause || "-",
    //             clause?.crby || "-",
    //             clause?.cron || "-"
    //           ])
    //         ];
      
    //         // Apply values to table
    //         table.values = values;
      
    //         await context.sync();
    //         console.log("✅ Table ", table, " inserted successfully.");
    //         // console.log("✅ Table inserted successfully.");

    //       });
    //     } catch (error) {
    //       console.error("❌ Error copying to Word:", error.message || error);
    //       // Optional: show error in UI instead of alert
    //     }
    //   }

    async function copyToWord(clauses) {
        try {
            await Word.run(async (context) => {
                const body = context.document.body;
                const rowCount = clauses.length + 1; // +1 for header row
                const colCount = 5;
    
                // Insert table title as a paragraph
                const titleParagraph = body.insertParagraph("Clause List", Word.InsertLocation.end);
                titleParagraph.font.size = 14; // Slightly larger font for the title
                titleParagraph.font.bold = true; // Make the title bold
                titleParagraph.alignment = Word.Alignment.centered; // Center the title
    
                // Add a small space between title and table (optional)
                body.insertParagraph("", Word.InsertLocation.end); // Empty paragraph for spacing
    
                // Insert table at the end of the document
                const table = body.insertTable(rowCount, colCount, Word.InsertLocation.end);
                table.style = "Grid Table 4 - Accent 1";
                table.getRange().font.size = 12;
                table.getRange().font.bold = false; // Explicitly set table content to not bold
    
                // Debug: log clauses to verify content
                console.log("Clauses to insert:", clauses);
    
                // Prepare table data (headers + rows)
                const values = [
                    ["Clause ID", "Title", "Description", "Created By", "Created On"],
                    ...clauses.map(clause => [
                        clause?.id || "-",
                        clause?.causetitle || "-",
                        clause?.cause || "-",
                        clause?.crby || "-",
                        clause?.cron || "-"
                    ])
                ];
    
                // Apply values to table
                table.values = values;
    
                await context.sync();
                console.log("✅ Table ", table, " inserted successfully.");
            });
        } catch (error) {
            console.error("❌ Error copying to Word:", error.message || error);
        }
    }
    // Simple Format
    // async function copyToWord(clauses) {
    //     try {
    //         await Word.run(async (context) => {
    //             console.log("Copying to Word:", clauses);
    //             const body = context.document.body;
    
    //             // Insert header as a paragraph
    //             body.insertParagraph("Clause ID | Title | Description | Created By | Created On", Word.InsertLocation.end);
    //             body.insertParagraph("------------------------------------------------------------", Word.InsertLocation.end);
    
    //             // Insert each clause as a paragraph
    //             clauses.forEach(clause => {
    //                 const clauseText = `${clause.id || '-'} | ${clause.causetitle || '-'} | ${clause.cause || '-'} | ${clause.crby || '-'} | ${clause.cron || '-'}`;
    //                 const paragraph = body.insertParagraph(clauseText, Word.InsertLocation.end);
    //                 paragraph.font.size = 10;
    //             });
    
    //             // Add some spacing at the end
    //             body.insertParagraph("", Word.InsertLocation.end);
    
    //             await context.sync();
    
    //             console.log("Data inserted into Word document");
    //         });
    //     } catch (error) {
    //         console.error("Error copying to Word:", error);
    //         alert("Error copying to Word document: " + error.message);
    //     }
    // }

    departmentSelect.addEventListener("change", (e) => {
        const deptId = e.target.value;
        if (deptId) {
            fetchAgreementTypeList(deptId);
        } else {
            agreementTypeSelect.innerHTML = `<option value="">Select a department first</option>`;
            agreementTypeSelect.disabled = true;
            proceedBtn.disabled = true;
        }
    });

    agreementTypeSelect.addEventListener("change", updateProceedButtonState);

    proceedBtn.addEventListener("click", async () => {
        const deptId = departmentSelect.value;
        const agTypeId = agreementTypeSelect.value;
        const apiUrl = "https://addinapi.convergelego.com/api/CompanyMaster/GetMstCauseLisit";
        const copyBtn = document.getElementById("copyToWordBtn");

        showLoader();
        copyBtn.style.display = "none"; // Hide button before fetching new data
        
        try {
            const response = await fetch(apiUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Userid": UserId,
                    "Key": Token,
                    "Token": Token,
                    "Comcode": ComCode
                },
                body: JSON.stringify({
                    deptid: deptId,
                    agrid: agTypeId,
                    statusid: "1"
                })
            });

            const result = await response.json();
            renderClausesTable(result?.Detail?.data || []);
        } catch (error) {
            console.error("Clause fetch error:", error);
            renderClausesTable([]);
        } finally {
            hideLoader();
        }
    });

    function renderClausesTable(clauses) {
        const container = document.getElementById("clausesTableContainer");
        const tbody = document.getElementById("clausesTableBody");
        const copyBtn = document.getElementById("copyToWordBtn");

        tbody.innerHTML = "";

        if (!clauses.length) {
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;">No clauses found.</td></tr>`;
            container.style.display = "block";
            copyBtn.style.display = "none";
            return;
        }

        clauses.forEach(c => {
            const row = document.createElement("tr");
            row.innerHTML = `
                <td>${c.id || '-'}</td>
                <td>${c.causetitle || '-'}</td>
                <td style="white-space: pre-wrap; max-width: 300px;">${c.cause || '-'}</td>
                <td>${c.crby || '-'}</td>
                <td>${c.cron || '-'}</td>
            `;
            tbody.appendChild(row);
        });

        container.style.display = "block";
        copyBtn.style.display = "block"; // Show button when data exists

        // Remove any existing listeners to prevent multiple bindings
        copyBtn.removeEventListener("click", copyBtn._listener);
        
        // Add new click listener
        const clickListener = () => copyToWord(clauses);
        copyBtn._listener = clickListener;
        copyBtn.addEventListener("click", clickListener);
    }

    // Initial fetch
    await fetchDepartmentList();
});