console.log("Before Office.onReady");

Office.onReady(() => {
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

    // Display login data in the userInfo section (uncomment if needed)
    // userInfoDiv.innerHTML = `
    //     <p><strong>Login ID:</strong> ${loginData.loginid}</p>
    //     <p><strong>Employee Name:</strong> ${loginData.Empname}</p>
    //     <p><strong>Company Code:</strong> ${loginData.comcode}</p>
    //     <p><strong>Company Name:</strong> ${loginData.CompanyName}</p>
    //     <p><strong>Department ID:</strong> ${loginData.Depid}</p>
    //     <p><strong>User Type:</strong> ${loginData.Utype}</p>
    // `;

    // Fetch department list from API
    const apiUrl = `https://api.convergelego.com/api/AddLegalAgreement/Departmentlist?companycode=${ComCode}&status=${status}`;

    fetch(apiUrl, {
        method: "GET",
        headers: {
            "Userid": UserId,
            "Key": "0",
            "Comcode": ComCode
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log("Department list response:", data);

        // Clear "Loading..." option
        departmentSelect.innerHTML = "";

        // Check if data is an array and has items
        if (Array.isArray(data) && data.length > 0) {
            // Add a default "Select Department" option
            const defaultOption = document.createElement("option");
            defaultOption.value = "";
            defaultOption.textContent = "Select Department";
            departmentSelect.appendChild(defaultOption);

            // Populate dropdown with department list
            data.forEach(department => {
                const option = document.createElement("option");
                option.value = department.Depid; // Assuming Depid is the department ID
                option.textContent = department.Depname || department.Depid; // Use Depname if available, else Depid
                departmentSelect.appendChild(option);
            });

            // Optionally pre-select the user's department if available in loginData
            if (loginData.Depid) {
                departmentSelect.value = loginData.Depid;
            }
        } else {
            departmentSelect.innerHTML = `<option value="">No departments found</option>`;
        }
    })
    .catch(error => {
        console.error("Error fetching department list:", error);
        departmentSelect.innerHTML = `<option value="">Error loading departments</option>`;
    });

    // Populate agreementType dropdown (placeholder for now)
    agreementTypeSelect.innerHTML = ""; // Clear "Loading..." option
    agreementTypeSelect.innerHTML = `<option value="">Select Agreement Type</option>`;
    // Add API call for agreement types if needed
});