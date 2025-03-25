function generateAI() {
    const idea = document.getElementById("aiInput").value;
    const output = "Sample AI clause based on input: " + idea;
    document.getElementById("aiResult").innerText = output;
  }
  
  function insertAIClause() {
    const text = document.getElementById("aiResult").innerText;
    if (!text) return alert("No AI clause to insert.");
  
    Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text + "\n\n", Word.InsertLocation.end);
      await context.sync();
      document.getElementById("status").innerText = "AI Clause inserted!";
    }).catch(console.error);
  }
  