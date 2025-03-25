console.log("Before Office.onReady");

Office.onReady(async () => {
    console.log("Office.onReady triggered");

    // Retrieve login data from localStorage
    if (!localStorage.getItem("loginData")) {
        console.log("No login data found");
        window.location.href = "index.html"; // Redirect to login if no data
        return;
    }

    const loginData = JSON.parse(localStorage.getItem("loginData"));
    const UserId = loginData.loginid;
    const ComCode = loginData.comcode;
    const status = loginData.status;
    console.log("Login data:", loginData);

    const userInfoDiv = document.getElementById("userInfo");
    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");

    // Check if login data exists
    if (!loginData) {
        userInfoDiv.innerHTML = "<p>Please log in first.</p>";
        window.location.href = "index.html"; // Redirect to login if no data
        return;
    }

    // Function to fetch and populate department list
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

            // Clear "Loading..." option
            departmentSelect.innerHTML = "";

            // Check if result.status == 200
            if (result.status == 200) {
                // Add a default "Select Department" option
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Department";
                departmentSelect.appendChild(defaultOption);

                // Assuming the department list is in result.Detail.data
                const departments = result?.Detail?.data || [];
                if (Array.isArray(departments) && departments.length > 0) {
                    // Populate dropdown with department list
                    departments.forEach(department => {
                        const option = document.createElement("option");
                        option.value = department.DeptId; // Use DeptId as the value
                        option.textContent = department.DeptName || department.DeptId; // Use DeptName if available, else DeptId
                        departmentSelect.appendChild(option);
                    });

                    // Optionally pre-select the user's department if available in loginData
                    if (loginData.Depid) {
                        departmentSelect.value = loginData.Depid;
                        // Trigger the agreement type fetch if a department is pre-selected
                        fetchAgreementTypeList(loginData.Depid);
                    }
                } else {
                    departmentSelect.innerHTML = `<option value="">No departments found</option>`;
                }
            } else {
                departmentSelect.innerHTML = `<option value="">No departments found</option>`;
            }
        } catch (error) {
            console.error("Error fetching department list:", error);
            departmentSelect.innerHTML = `<option value="">No departments found</option>`;
        }
    }

    // Function to fetch and populate agreement type list based on selected department
    async function fetchAgreementTypeList(deptId) {
        const apiUrl = `https://lapi.convergelego.com/api/STD/AgreementType/MstAgTypelist?comcode=${ComCode}&depid=${deptId}`;

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
            console.log("Agreement type list response:", result);

            // Clear "Loading..." option
            agreementTypeSelect.innerHTML = "";

            // Check if result.status == 200
            if (result.status == 200) {
                // Add a default "Select Agreement Type" option
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Agreement Type";
                agreementTypeSelect.appendChild(defaultOption);

                // Assuming the agreement type list is in result.Detail.data
                const agreementTypes = result?.Detail?.data || [];
                if (Array.isArray(agreementTypes) && agreementTypes.length > 0) {
                    // Populate dropdown with agreement type list
                    agreementTypes.forEach(agreement => {
                        const option = document.createElement("option");
                        option.value = agreement.AgTypeId; // Use AgTypeId as the value (adjust based on actual response)
                        option.textContent = agreement.AgTypeName || agreement.AgTypeId; // Use AgTypeName if available, else AgTypeId
                        agreementTypeSelect.appendChild(option);
                    });
                } else {
                    agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
                }
            } else {
                agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
            }
        } catch (error) {
            console.error("Error fetching agreement type list:", error);
            agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
        }
    }

    // Fetch department list on page load
    await fetchDepartmentList();

    // Add event listener to department dropdown to fetch agreement types when a department is selected
    departmentSelect.addEventListener("change", (event) => {
        const selectedDeptId = event.target.value;
        if (selectedDeptId) {
            fetchAgreementTypeList(selectedDeptId);
        } else {
            // Clear agreement type dropdown if no department is selected
            agreementTypeSelect.innerHTML = `<option value="">Select Agreement Type</option>`;
        }
    });
});