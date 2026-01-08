# Structural Report Generator (Prototype)

An **Express.js–based prototype application** designed to generate **structural engineering reports** using pre-defined document templates. The app automates report creation by injecting user-provided data into standardized templates, significantly reducing manual formatting effort.

This project currently serves as a **proof of concept** and will be reimplemented as a **Spring Boot–based application** in the coming months for improved scalability, maintainability, and production readiness.

---

## Features

- Generate structured reports from predefined templates
- Dynamic template editing using placeholders
- File upload support for template and data inputs
- Automated document generation workflow

---

## Tech Stack

- **Backend:** Node.js, Express
- **Template Processing:** Docxtemplater
- **File Uploads:** Multer
- **File Compression / Processing:** PizZip
- **Runtime:** JavaScript (Node.js)

---

## Key Libraries

- **[Express](https://expressjs.com/):** Server framework for handling routes and requests  
- **[Docxtemplater](https://docxtemplater.com/):** Injects dynamic data into `.docx` templates  
- **[Multer](https://github.com/expressjs/multer):** Handles multipart file uploads  
- **[PizZip](https://github.com/open-xml-templating/pizzip):** Processes zipped document formats such as `.docx`

---

## Current Status

**Prototype Stage**

- Core document generation workflow implemented
- Basic upload and processing pipeline complete
- Not production-ready
- Limited error handling and validation

This version was built to validate the overall approach and workflow before migrating to a more robust backend architecture.

---

## Future Plans

- Rewrite backend using **Spring Boot**
- Improve validation and error handling
- Add authentication and user management
- Expand template support and customization
- Improve performance and scalability
- Add persistence layer (database support)

---

## Motivation

Structural reports often follow strict formatting and structure. This project aims to **reduce repetitive manual work** by automating report generation while maintaining consistency and accuracy.

---

## Getting Started (Prototype)

```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
npm install
npm start
