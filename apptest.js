require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { imageSize } = require("image-size");
const multer = require("multer");

const app = express();
const PORT = 3030;
const templatePath = path.resolve(__dirname, "mpTemplate.docx");
const imageLabelMap = {
    image1: "Cover Image",
    image2: "Caisson Installation Drilling",
    image3: "Caisson Rebar Cage",
    image4: "Rebar Length",
    image25: "Tie Spacing",
    image26: "Tie Spacing (Additional View)",
    image6: "(28) 25M Vertical Rebar",
    image7: "(28) 25M Vertical Rebar (Additional View)",
    image8: "Anchor Rod Cage",
    image9: "Reinforcement Cage Lift",
    image10: "Caisson Concrete Cover",
    image11: "Caisson Concrete Cover (Additional View)",
    image12: "Caisson Concrete Pouring",
    image13: "WIC Pad Length",
    image14: "WIC Pad Rebar",
    image15: "Foundation Construction",
    image16: "Completed Foundation Construction",
};

app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const upload = multer({
    dest: "uploads/",
    limits: { fileSize: 10 * 1024 * 1024 },
});

function getTemplateTags() {
    if (!fs.existsSync(templatePath)) {
        throw new Error("Template file not found.");
    }
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    const text = doc.getFullText();
    const matches = text.match(/{\s*([^#\/^][^}]*)\s*}/g) || [];
    return [
        ...new Set(
            matches
                .map((tag) => tag.replace(/[{}]/g, "").trim())
                .map((tag) => tag.replace(/^%+/, ""))
                .filter(Boolean)
        ),
    ];
}

function inferFieldType(tag) {
    const normalized = tag.toLowerCase();
    if (normalized.includes("image") || normalized.includes("photo") || normalized.includes("pic")) {
        return "image";
    }
    if (normalized.includes("date")) {
        return "date";
    }
    if (normalized.includes("desc") || normalized.includes("address") || normalized.includes("notes")) {
        return "textarea";
    }
    return "text";
}

function toLabel(tag) {
    if (imageLabelMap[tag]) {
        return imageLabelMap[tag];
    }
    return tag
        .replace(/[_-]+/g, " ")
        .replace(/([a-z])([A-Z])/g, "$1 $2")
        .replace(/\b\w/g, (char) => char.toUpperCase());
}

function cleanupUploads(files = []) {
    files.forEach((file) => {
        fs.unlink(file.path, () => {});
    });
}

function shouldAutofillField(field) {
    if (!field || field.type !== "text" && field.type !== "textarea") {
        return false;
    }
    const normalized = field.name.toLowerCase();
    const structuredFields = ["sitecode", "sitename", "reportdate", "latn", "latw", "groundele", "pourdate"];
    return !structuredFields.includes(normalized);
}

function extractJsonObject(content) {
    if (!content || typeof content !== "string") return null;
    const fenceMatch = content.match(/```(?:json)?\s*([\s\S]*?)```/i);
    const jsonCandidate = fenceMatch ? fenceMatch[1] : content;
    const startIdx = jsonCandidate.indexOf("{");
    const endIdx = jsonCandidate.lastIndexOf("}");
    if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx) return null;
    try {
        return JSON.parse(jsonCandidate.slice(startIdx, endIdx + 1));
    } catch {
        return null;
    }
}

async function generateFieldSuggestions({ overview, fields, currentValues }) {
    const apiKey = (process.env.GROQ_KEY || "").trim();
    if (!apiKey) {
        throw new Error("GROQ_KEY is not set on the server.");
    }
    const model = (process.env.GROQ_MODEL || "llama-3.1-8b-instant").trim();
    const targets = fields.filter(shouldAutofillField).map((field) => ({
        name: field.name,
        label: field.label,
    }));

    const prompt = [
        "You are assisting with drafting engineering report form text.",
        "Using the project overview and any current field values, generate concise professional content.",
        "Return only JSON with this shape: {\"suggestions\":{\"fieldName\":\"value\"}}.",
        "Only include fields from the provided targets.",
        "If a field already has non-empty content, improve it but keep intent.",
        "",
        `Project overview: ${overview}`,
        `Targets: ${JSON.stringify(targets)}`,
        `Current values: ${JSON.stringify(currentValues)}`,
    ].join("\n");

    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: {
            Authorization: `Bearer ${apiKey}`,
            "Content-Type": "application/json",
        },
        body: JSON.stringify({
            model,
            temperature: 0.3,
            messages: [
                { role: "system", content: "You write clear and factual construction/inspection report content." },
                { role: "user", content: prompt },
            ],
        }),
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`LLM request failed (Groq): ${response.status} ${errorText}`);
    }

    const data = await response.json();
    const message = data?.choices?.[0]?.message?.content || "";
    const parsed = extractJsonObject(message);
    if (!parsed || typeof parsed !== "object" || typeof parsed.suggestions !== "object") {
        throw new Error("LLM response could not be parsed into suggestions JSON.");
    }

    const allowedNames = new Set(targets.map((item) => item.name));
    const cleaned = {};
    Object.entries(parsed.suggestions).forEach(([name, value]) => {
        if (allowedNames.has(name) && typeof value === "string") {
            cleaned[name] = value.trim();
        }
    });
    return cleaned;
}

app.get("/api/template-fields", (req, res) => {
    try {
        const fields = getTemplateTags().map((name) => ({
            name,
            type: inferFieldType(name),
            label: toLabel(name),
            section: inferFieldType(name) === "image" && name !== "image1" ? "photos_shared_by_contractor" : "general",
        }));
        res.json({ fields });
    } catch (error) {
        console.error("Failed to parse template:", error);
        res.status(500).json({ error: error.message });
    }
});

app.post("/api/ai-fill", async (req, res) => {
    try {
        const overview = (req.body?.overview || "").trim();
        const currentValues = req.body?.currentValues || {};
        if (!overview) {
            return res.status(400).json({ error: "Project overview is required." });
        }

        const fields = getTemplateTags().map((name) => ({
            name,
            type: inferFieldType(name),
            label: toLabel(name),
            section: inferFieldType(name) === "image" && name !== "image1" ? "photos_shared_by_contractor" : "general",
        }));

        const suggestions = await generateFieldSuggestions({ overview, fields, currentValues });
        return res.json({ suggestions });
    } catch (error) {
        console.error("AI autofill failed:", error);
        return res.status(500).json({ error: error.message });
    }
});

app.post("/generate", upload.any(), (req, res) => {
    try {
        const content = fs.readFileSync(templatePath, "binary");
        const zip = new PizZip(content);
        const filesByField = {};
        (req.files || []).forEach((file) => {
            filesByField[file.fieldname] = file;
        });

        const imageModule = new ImageModule({
            centered: false,
            fileType: "docx",
            getImage: (tagValue) => {
                if (!tagValue || !fs.existsSync(tagValue)) {
                    return Buffer.alloc(0);
                }
                return fs.readFileSync(tagValue);
            },
            getSize: (img) => {
                try {
                    if (!img.length) return [1, 1];
                    const size = imageSize(img);
                    const maxWidth = 600;
                    if (size.width <= maxWidth) return [size.width, size.height];
                    const ratio = maxWidth / size.width;
                    return [maxWidth, Math.round(size.height * ratio)];
                } catch {
                    return [400, 300];
                }
            },
        });

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            modules: [imageModule],
        });

        const templateTags = getTemplateTags();
        const data = {};
        templateTags.forEach((tag) => {
            if (filesByField[tag]) {
                data[tag] = filesByField[tag].path;
            } else {
                data[tag] = req.body[tag] || "";
            }
        });

        doc.setData(data);
        doc.render();

        const buffer = doc.getZip().generate({ type: "nodebuffer" });
        const reportDate = data.reportdate || new Date().toISOString().split("T")[0];
        const siteName = data.sitename || "report";

        res.set({
            "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "Content-Disposition": `attachment; filename="${siteName}_${reportDate}.docx"`,
            "Content-Length": buffer.length,
        });
        res.send(buffer);
        cleanupUploads(req.files);
    } catch (error) {
        cleanupUploads(req.files);
        console.error("Document generation failed:", error);
        res.status(500).send(`Failed to generate document: ${error.message}`);
    }
});

app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
        return res.status(400).send("File too large. Maximum size is 10MB.");
    }
    console.error("Unhandled upload error:", error);
    return res.status(500).send(`File upload error: ${error.message}`);
});

app.get("/health", (req, res) => {
    const groqKeyPresent = !!(process.env.GROQ_KEY && String(process.env.GROQ_KEY).trim());
    res.json({
        status: "OK",
        timestamp: new Date().toISOString(),
        aiAutofillProvider: "groq",
        groqKeyConfigured: groqKeyPresent,
    });
});

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log(`Health check available at http://localhost:${PORT}/health`);
    console.log(
        "AI auto-fill uses Groq only (OPENAI/Gemini URLs are not called). Require GROQ_KEY; optional GROQ_MODEL defaults to llama-3.1-8b-instant."
    );
});