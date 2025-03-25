// Office.onReady(() => {
//     const params = new URLSearchParams(window.location.search);
//     const type = params.get("type");
  
//     document.getElementById("loginForm").onsubmit = async function (e) {
//         e.preventDefault();
    
//         const username = document.getElementById("username").value;
//         const password = document.getElementById("password").value;
    
//         if (!username || !password) {
//           alert("Please enter username and password.");
//           return;
//         }
    
//         try {
//           const apiUrl = `https://addinapi.convergelego.com/api/Login/LoginCheck?pwd=${encodeURIComponent(password)}`;
    
//           const res = await fetch(apiUrl, {
//             method: "GET",
//             headers: {
//               "Userid": username,
//               "Key": "0"
//             }
//           });
    
//           const result = await res.json();
    
//           if (res.ok && result.status === true) {
//             showContent(type);
//           } else {
//             alert(result.message || "Invalid credentials");
//           }
    
//         } catch (err) {
//           console.error("Login API Error:", err);
//           alert("Login failed. Please check your network or credentials.");
//         }
//       };
//   });
  
//   function showContent(type) {
//     document.getElementById("loginForm").style.display = "none";
//     const content = document.getElementById("content");
//     const title = document.getElementById("contentTitle");
//     const input = document.getElementById("clauseInput");
  
//     content.style.display = "block";
//     title.textContent = type === "ai" ? "AI Clauses Generator" : "Manual Clauses Entry";
//     input.placeholder = type === "ai" ? "Generate or enter AI-based clause" : "Enter your clause manually";
//   }
  
//   function insertClause() {
//     const text = document.getElementById("clauseInput").value;
//     if (!text) return alert("Please enter a clause.");
  
//     Word.run(async (context) => {
//       const range = context.document.getSelection();
//       range.insertText(text + "\n\n", Word.InsertLocation.end);
//       await context.sync();
//       document.getElementById("status").textContent = "Clause inserted successfully!";
//     }).catch((err) => {
//       console.error("Error inserting clause:", err);
//     });
//   }
  
Office.onReady(() => {
    // Get the type parameter from URL
    const params = new URLSearchParams(window.location.search);
    const type = params.get("type"); // 'clauses' or 'ai'

    document.getElementById("loginForm").onsubmit = async function (e) {
        e.preventDefault();

        const username = document.getElementById("username").value.trim();
        const password = document.getElementById("password").value.trim();
        const statusDiv = document.getElementById("status");

        if (!username || !password) {
            statusDiv.innerText = "Please enter username and password.";
            return;
        }

        try {
            const apiUrl = `https://addinapi.convergelego.com/api/Login/LoginCheck?pwd=${encodeURIComponent(password)}`;

            const res = await fetch(apiUrl, {
                method: "GET",
                headers: {
                    "Userid": username,
                    "Key": "0"
                }
            });

            const result = await res.json();
            console.log("Login result:", result);
            if (result.status === 200) {
                // Successful login, navigate based on type
                console.log("Login successful. Redirecting...", type);
                localStorage.setItem("loginData", result.Detail.data);
                if (type === "clauses") {
                    window.location.href = "clauses.html";
                } else if (type === "ai") {
                    window.location.href = "aiclauses.html";
                } else {
                    statusDiv.innerText = "Invalid page type.";
                }
            } else {
                statusDiv.innerText = result.message || "Invalid credentials";
            }

        } catch (err) {
            console.error("Login error:", err);
            statusDiv.innerText = "Login failed. Please try again.";
        }
    };
});