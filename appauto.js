const fs = require("fs");
const path = require("path");
const express = require("express");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const PORT = 3000;

app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
    try {
        const templatePath = path.join(__dirname, "mpTemplate.docx");

        // Confirm file exists
        if (!fs.existsSync(templatePath)) {
            return res.status(404).send("template.docx not found.");
        }

        const content = fs.readFileSync(templatePath, "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip);

        const rawText = doc.getFullText();

        // Match { tag } style tags
        const matches = rawText.match(/{\s*([^}]+?)\s*}/g) || [];
        const uniqueTags = [...new Set(matches.map(tag => tag.replace(/[{}]/g, "").trim()))];

        // HTML Form generation
        let formHtml = `
            <html>
            <head>
                <title>Docx Form Generator</title>
            </head>
            <body>
                <h1>Fill out the form</h1>
                <form method="POST" action="/generate">
        `;

        uniqueTags.forEach(tag => {
            formHtml += `
                <label for="${tag}">${tag}:</label><br>
                <input type="text" name="${tag}" style="width: 300px;" required><br><br>
            `;
        });

        formHtml += `
                <button type="submit">Generate DOCX</button>
                </form>
            </body>
            </html>
        `;

        res.send(formHtml);
    } catch (err) {
        console.error("Error:", err);
        res.status(500).send("Internal server error");
    }
});

app.listen(PORT, () => {
    console.log(`✅ Server running at http://localhost:${PORT}`);
});
