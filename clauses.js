function insertClause() {
    const text = document.getElementById("clauseInput").value;
    if (!text) return alert("Please enter a clause.");
  
    Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text + "\n\n", Word.InsertLocation.end);
      await context.sync();
      document.getElementById("status").innerText = "Clause inserted successfully!";
    }).catch(console.error);
  }
  