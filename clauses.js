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

    async function copyToWord(clause) {
        try {
            await Word.run(async (context) => {
                const body = context.document.body;

                const title = clause?.causetitle || "Untitled Clause";
                const desc = clause?.cause || "-";

                const titlePara = body.insertParagraph("", Word.InsertLocation.end);
                const titleRange = titlePara.insertText("Title - ", Word.InsertLocation.start);
                titleRange.font.bold = true;
                titleRange.font.size = 14;

                const titleTextRange = titlePara.insertText(title, Word.InsertLocation.end);
                titleTextRange.font.bold = false;
                titleTextRange.font.size = 14;

                const descLabelPara = body.insertParagraph("", Word.InsertLocation.end);
                const descLabelRange = descLabelPara.insertText("Description -", Word.InsertLocation.start);
                descLabelRange.font.bold = true;
                descLabelRange.font.size = 14;

                const descPara = body.insertParagraph(desc, Word.InsertLocation.end);
                descPara.font.size = 14;
                descPara.font.bold = false;
                descPara.spacingAfter = 20;

                await context.sync();
                console.log("✅ Inserted:", title);
            });
        } catch (error) {
            console.error("❌ Error inserting clause:", error.message || error);
        }
    }

    proceedBtn.addEventListener("click", async () => {
        const deptId = departmentSelect.value;
        const agTypeId = agreementTypeSelect.value;
        const apiUrl = "https://addinapi.convergelego.com/api/CompanyMaster/GetMstCauseLisit";

        const container = document.getElementById("clausesCardContainer");

        showLoader();
        container.style.display = "none";

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
        const container = document.getElementById("clausesCardContainer");
        container.innerHTML = "";

        if (!clauses.length) {
            container.innerHTML = `<p style="text-align:center;">No clauses found.</p>`;
            container.style.display = "block";
            return;
        }

        clauses.forEach((clause, index) => {
            const card = document.createElement("div");
            card.className = "card";

            const copyBtn = document.createElement("button");
            copyBtn.className = "copy-btn";
            copyBtn.innerText = "Copy";
            copyBtn.addEventListener("click", () => copyToWord(clause));

            const titleSpan = document.createElement("span");
            titleSpan.className = "title-span";
            titleSpan.innerText = clause.causetitle || "-";

            const descSpan = document.createElement("span");
            descSpan.className = "description-span";
            descSpan.innerText = clause.cause || "-";

            card.appendChild(copyBtn);
            card.appendChild(titleSpan);
            card.appendChild(descSpan);
            container.appendChild(card);
        });

        container.style.display = "block";
    }

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

    // Initial fetch
    await fetchDepartmentList();
});
