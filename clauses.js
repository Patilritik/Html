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

    console.log("Login data:", loginData);

    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");
    const proceedBtn = document.getElementById("proceedBtn");

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
            console.log("Department list response:", result);
            departmentSelect.innerHTML = "";

            if (result.status == 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Department";
                departmentSelect.appendChild(defaultOption);

                const departments = result?.Detail?.data || [];
                departments.forEach(department => {
                    const option = document.createElement("option");
                    option.value = department.DeptId;
                    option.textContent = department.DeptName || department.DeptId;
                    departmentSelect.appendChild(option);
                });

                if (loginData.Depid) {
                    departmentSelect.value = loginData.Depid;
                    fetchAgreementTypeList(loginData.Depid);
                }
            } else {
                departmentSelect.innerHTML = `<option value="">No departments found</option>`;
            }
        } catch (error) {
            console.error("Error fetching department list:", error);
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
            console.log("Agreement type list:", result);
            agreementTypeSelect.innerHTML = "";

            if (result.status == 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Agreement Type";
                agreementTypeSelect.appendChild(defaultOption);

                const agreementTypes = result?.Detail?.data || [];
                agreementTypes.forEach(agreement => {
                    const option = document.createElement("option");
                    option.value = agreement.Agtypeid;
                    option.textContent = agreement.AgtypeDesc || agreement.Agtypeid;
                    agreementTypeSelect.appendChild(option);
                });
            } else {
                agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
            }
        } catch (error) {
            console.error("Error fetching agreement type list:", error);
            agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
        }

        updateProceedButtonState();
    }

    function updateProceedButtonState() {
        const deptSelected = departmentSelect.value !== "";
        const agreementSelected = agreementTypeSelect.value !== "";
        proceedBtn.disabled = !(deptSelected && agreementSelected);
    }

    await fetchDepartmentList();

    departmentSelect.addEventListener("change", (event) => {
        const selectedDeptId = event.target.value;
        if (selectedDeptId) {
            fetchAgreementTypeList(selectedDeptId);
        } else {
            agreementTypeSelect.innerHTML = `<option value="">Select a department first</option>`;
            agreementTypeSelect.disabled = true;
            proceedBtn.disabled = true;
        }
    });

    agreementTypeSelect.addEventListener("change", () => {
        updateProceedButtonState();
    });

    proceedBtn.addEventListener("click", async () => {
        const selectedDeptId = departmentSelect.value;
        const selectedAgreementId = agreementTypeSelect.value;

        console.log(`Proceeding with Department ID: ${selectedDeptId}, Agreement Type ID: ${selectedAgreementId}`);

        const apiUrl = "https://addinapi.convergelego.com/api/CompanyMaster/GetMstCauseLisit";

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
                    deptid: selectedDeptId,
                    agrid: selectedAgreementId,
                    statusid: "1"
                })
            });

            const result = await response.json();
            console.log("Clauses API response:", result);

            renderClausesTable(result?.Detail?.data || []);
        } catch (error) {
            console.error("Error fetching clauses:", error);
            renderClausesTable([]);
        }
    });

    function renderClausesTable(clauses) {
        const container = document.getElementById("clausesTableContainer");
        const tbody = document.getElementById("clausesTableBody");

        tbody.innerHTML = "";

        if (!clauses.length) {
            tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;">No clauses found.</td></tr>`;
            container.style.display = "block";
            return;
        }

        clauses.forEach(clause => {
            const row = document.createElement("tr");

            row.innerHTML = `
                <td>${clause.id || '-'}</td>
                <td>${clause.causetitle || '-'}</td>
                <td style="white-space: pre-wrap; max-width: 300px;">${clause.cause || '-'}</td>
                <td>${clause.statusdesc || '-'}</td>
                <td>${clause.crby || '-'}</td>
                <td>${clause.cron || '-'}</td>
            `;

            tbody.appendChild(row);
        });

        container.style.display = "block";
    }
});
