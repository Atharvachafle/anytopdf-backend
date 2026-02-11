const express = require('express');
const multer = require('multer');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const puppeteer = require('puppeteer');
const PptxGenJS = require("pptxgenjs");
const { PDFDocument } = require('pdf-lib');

// Import image-size (Handle both newer and older versions)
const imageSizePkg = require('image-size');
const sizeOf = typeof imageSizePkg.imageSize === 'function' ? imageSizePkg.imageSize : imageSizePkg;

const app = express();
const port = 3000;

// ==========================================
// ⚙️ PATH CONFIGURATION (Auto-detects Windows vs Linux)
// ==========================================
const isWindows = process.platform === 'win32';

const GS_PATH = isWindows 
    ? '"C:\\Program Files\\gs\\gs10.06.0\\bin\\gswin64c.exe"' 
    : 'gs'; 

const SOFFICE_PATH = isWindows 
    ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"' 
    : 'soffice';

const POPPLER_PATH = isWindows 
    ? '"C:\\Program Files\\poppler\\Library\\bin\\pdftocairo.exe"' 
    : 'pdftocairo';

const QPDF_PATH = isWindows 
    ? '"C:\\Program Files\\qpdf\\bin\\qpdf.exe"' 
    : 'qpdf';


    
// ==========================================
// 🚀 MIDDLEWARE
// ==========================================
// Allow all requests (CORS FIX)
app.use(cors({
    origin: '*', 
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type']
}));
app.use(express.json());

// ==========================================
// 📂 FOLDER SETUP
// ==========================================
const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'outputs');
const tempDir = path.join(__dirname, 'temp_profiles');

// Ensure all directories exist
[uploadDir, outputDir, tempDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
        console.log(`📂 Creating folder: ${dir}`);
        fs.mkdirSync(dir, { recursive: true });
    }
});

const upload = multer({ dest: 'uploads/' });

// =========================================================================
// 🛣️ ROUTES
// =========================================================================

// 0. HEALTH CHECK
app.get('/', (req, res) => {
    res.send('✅ PDF Server is Running! You can now use the tool.');
});

// 1️⃣ COMPRESS PDF (Calculates Size & Returns JSON)
app.post('/compress-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    
    const inputPath = path.resolve(req.file.path);
    const fileID = `compressed-${Date.now()}.pdf`; 
    const outputPath = path.join(outputDir, fileID);
    const level = req.body.level || 'recommended';

    let qualitySetting = '/ebook'; 
    if (level === 'extreme') qualitySetting = '/screen';
    if (level === 'less') qualitySetting = '/printer';

    console.log(`📉 Processing: ${req.file.originalname}`);

    const command = `${GS_PATH} -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${qualitySetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${outputPath}" "${inputPath}"`;

    exec(command, (error) => {
        try { fs.unlinkSync(inputPath); } catch(e) {}

        if (error) {
            console.error("GS Error:", error);
            return res.status(500).json({ error: 'Compression failed' });
        }

        if (fs.existsSync(outputPath)) {
            const stats = fs.statSync(outputPath);
            const sizeInMB = (stats.size / 1024 / 1024).toFixed(2);

            console.log(`✅ Success! New Size: ${sizeInMB} MB`);
            
            res.json({ 
                success: true, 
                downloadUrl: `http://localhost:${port}/download/${fileID}`,
                newSize: sizeInMB
            });
        } else {
            res.status(500).json({ error: 'Output file missing' });
        }
    });
});

// 2️⃣ DOWNLOAD ROUTE
app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(outputDir, filename);

    if (fs.existsSync(filePath)) {
        res.download(filePath, 'compressed.pdf', (err) => {
            if (err) console.error("Download Error:", err);
            // Optional: Delete file after download
            // try { fs.unlinkSync(filePath); } catch(e) {} 
        });
    } else {
        res.status(404).send("File not found or expired.");
    }
});

// =========================================================================
// 3️⃣ EXCEL TO PDF ROUTE (Fixed Extension)
// =========================================================================
app.post('/excel-to-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPath = path.resolve(req.file.path);
    // 1. FIX: Add extension so LibreOffice detects it as Excel
    const originalExt = path.extname(req.file.originalname);
    const inputPathWithExt = inputPath + originalExt;

    try {
        fs.renameSync(inputPath, inputPathWithExt);
    } catch (err) {
        console.error("Rename Error:", err);
        return res.status(500).send("File processing error.");
    }

    const uniqueProfileDir = path.join(tempDir, `profile-${Date.now()}`);
    const profileUrl = 'file:///' + uniqueProfileDir.replace(/\\/g, '/');

    console.log(`📊 Converting Excel to PDF: ${req.file.originalname}`);

    // 2. FIX: Use inputPathWithExt in the command
    const command = `${SOFFICE_PATH} -env:UserInstallation="${profileUrl}" --headless --convert-to pdf --outdir "${outputDir}" "${inputPathWithExt}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.rmSync(uniqueProfileDir, { recursive: true, force: true }); } catch(e) {}
        
        if (error) {
            console.error("Conversion Failed:", stderr);
            try { fs.unlinkSync(inputPathWithExt); } catch(e) {} // Cleanup
            return res.status(500).send('PDF creation failed.');
        }

        // LibreOffice uses the filename from inputPathWithExt
        const outputName = path.parse(inputPathWithExt).name + '.pdf';
        const generatedFile = path.join(outputDir, outputName);

        if (fs.existsSync(generatedFile)) {
            console.log("✅ Success! Sending PDF.");
            res.download(generatedFile, (err) => {
                if(err) console.error("Download Error:", err);
                try { 
                    fs.unlinkSync(inputPathWithExt); 
                    fs.unlinkSync(generatedFile); 
                } catch(e) {}
            });
        } else {
            console.error("Output file missing.");
            res.status(500).send('PDF file not found.');
        }
    });
});

// 4️⃣ HTML TO PDF ROUTE
app.post('/html-to-pdf', async (req, res) => {
    console.log("🔵 Request Received:", req.body);

    const { url, format = 'A4', orientation = 'portrait' } = req.body;

    if (!url) {
        console.error("❌ Error: No URL provided");
        return res.status(400).send('URL is missing.');
    }

    const targetUrl = url.startsWith('http') ? url : `https://${url}`;
    const outputFilename = `website-${Date.now()}.pdf`;
    const outputPath = path.join(outputDir, outputFilename);

    let browser;
    try {
        console.log(`🚀 Launching Browser for: ${targetUrl}`);
        
        browser = await puppeteer.launch({
            headless: "new",
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });

        const page = await browser.newPage();
        
        console.log("🌍 Navigate to page...");
        await page.goto(targetUrl, { waitUntil: 'networkidle2', timeout: 60000 });

        console.log("📄 Generating PDF...");
        await page.pdf({
            path: outputPath,
            format: format,
            landscape: orientation === 'landscape',
            printBackground: true,
            margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' }
        });

        console.log("✅ PDF Created at:", outputPath);

        res.download(outputPath, (err) => {
            if (err) console.error("❌ Download Error:", err);
            setTimeout(() => { try { fs.unlinkSync(outputPath); } catch(e){} }, 60000);
        });

    } catch (err) {
        console.error("❌ PUPPETEER ERROR:", err.message);
        res.status(500).send("Server Error: " + err.message);
    } finally {
        if (browser) await browser.close();
    }
});

// 5️⃣ PDF TO EXCEL (Python)
app.post('/convert-to-excel', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPath = path.resolve(req.file.path);
    const outputFilename = path.parse(req.file.originalname).name + '.xlsx';
    const outputPath = path.join(outputDir, outputFilename);
    
    const pythonScript = path.join(__dirname, 'convert_excelpte.py');

    console.log(`📊 Extracting Tables: ${req.file.originalname}`);

    const command = `python "${pythonScript}" "${inputPath}" "${outputPath}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.unlinkSync(inputPath); } catch(e) {}

        if (error || stdout.includes("ERROR")) {
            console.error(`Error: ${stdout || stderr}`);
            return res.status(500).send('Conversion failed.');
        }

        if (fs.existsSync(outputPath)) {
            console.log("✅ Tables Extracted Successfully!");
            res.download(outputPath, outputFilename, (err) => {
                try { fs.unlinkSync(outputPath); } catch(e) {}
            });
        } else {
            res.status(500).send('Output file missing.');
        }
    });
});

// 6️⃣ PDF TO POWERPOINT
app.post('/pdf-to-pptx', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPath = path.resolve(req.file.path);
    const outputFilename = `presentation-${Date.now()}.pptx`;
    const outputPath = path.join(outputDir, outputFilename);
    const imgPrefix = path.join(outputDir, `slide-${Date.now()}`);

    console.log(`🖼️ Processing: ${req.file.originalname}`);

    const command = `${POPPLER_PATH} -jpeg -r 150 "${inputPath}" "${imgPrefix}"`;

    exec(command, async (error) => {
        if (error) {
            console.error("Poppler Error:", error);
            return res.status(500).send('Image generation failed.');
        }

        try {
            const files = fs.readdirSync(outputDir);
            const slideImages = files
                .filter(f => f.startsWith(path.basename(imgPrefix)) && f.endsWith('.jpg'))
                .sort((a, b) => {
                    const numA = parseInt(a.match(/-(\d+)\.jpg$/)[1]);
                    const numB = parseInt(b.match(/-(\d+)\.jpg$/)[1]);
                    return numA - numB;
                });

            if (slideImages.length === 0) throw new Error("No images generated.");

            const firstImgPath = path.join(outputDir, slideImages[0]);
            const imgBuffer = fs.readFileSync(firstImgPath);
            const dimensions = sizeOf(imgBuffer);

            const widthInches = dimensions.width / 150;
            const heightInches = dimensions.height / 150;

            const pres = new PptxGenJS();
            pres.defineLayout({ name: 'PDF_SIZE', width: widthInches, height: heightInches });
            pres.layout = 'PDF_SIZE';

            slideImages.forEach(img => {
                const slide = pres.addSlide();
                const imgPath = path.join(outputDir, img);
                slide.addImage({ 
                    path: imgPath, x: 0, y: 0, w: widthInches, h: heightInches 
                });
            });

            await pres.writeFile({ fileName: outputPath });
            console.log("✅ PPTX Created Successfully!");

            res.download(outputPath, (err) => {
                try {
                    fs.unlinkSync(inputPath);
                    fs.unlinkSync(outputPath);
                    slideImages.forEach(img => fs.unlinkSync(path.join(outputDir, img)));
                } catch(e) {}
            });

        } catch (err) {
            console.error(err);
            res.status(500).send('PPTX creation failed.');
        }
    });
});

// 7️⃣ PDF TO WORD (Python)
app.post('/convert-to-word', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    const inputPath = path.resolve(req.file.path);
    const outputFilename = path.parse(req.file.originalname).name + '.docx';
    const outputPath = path.join(outputDir, outputFilename);
    
    const pythonScript = path.join(__dirname, 'convert_tablesptw.py');

    console.log(`📝 Converting (Table Mode): ${req.file.originalname}`);

    const command = `python "${pythonScript}" "${inputPath}" "${outputPath}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.unlinkSync(inputPath); } catch(e) {}

        if (error || stdout.includes("ERROR")) {
            console.error(`Error: ${stdout || stderr}`);
            return res.status(500).send('Conversion failed.');
        }

        if (fs.existsSync(outputPath)) {
            console.log("✅ Conversion Successful!");
            res.download(outputPath, outputFilename, (err) => {
                try { fs.unlinkSync(outputPath); } catch(e) {}
            });
        } else {
            res.status(500).send('Output file missing.');
        }
    });
});

// =========================================================================
// 8️⃣ POWERPOINT TO PDF ROUTE (Fixed Extension)
// =========================================================================
app.post('/pptx-to-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPath = path.resolve(req.file.path);
    // 1. FIX: Add extension so LibreOffice detects it as PowerPoint
    const originalExt = path.extname(req.file.originalname);
    const inputPathWithExt = inputPath + originalExt;

    try {
        fs.renameSync(inputPath, inputPathWithExt);
    } catch (err) {
        console.error("Rename Error:", err);
        return res.status(500).send("File processing error.");
    }
    
    const uniqueProfileDir = path.join(tempDir, `profile-${Date.now()}`);
    const profileUrl = 'file:///' + uniqueProfileDir.replace(/\\/g, '/');

    console.log(`📊 Converting PPTX to PDF: ${req.file.originalname}`);

    // 2. FIX: Use inputPathWithExt in the command
    const command = `${SOFFICE_PATH} -env:UserInstallation="${profileUrl}" --headless --convert-to pdf --outdir "${outputDir}" "${inputPathWithExt}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.rmSync(uniqueProfileDir, { recursive: true, force: true }); } catch(e) {}

        if (error) {
            console.error("Conversion Failed:", stderr);
            try { fs.unlinkSync(inputPathWithExt); } catch(e) {} // Cleanup
            return res.status(500).send('PDF creation failed.');
        }

        // LibreOffice uses the filename from inputPathWithExt
        const outputName = path.parse(inputPathWithExt).name + '.pdf';
        const generatedFile = path.join(outputDir, outputName);

        if (fs.existsSync(generatedFile)) {
            console.log("✅ Success! Sending PDF.");
            res.download(generatedFile, (err) => {
                if(err) console.error("Download Error:", err);
                try { 
                    fs.unlinkSync(inputPathWithExt); 
                    fs.unlinkSync(generatedFile); 
                } catch(e) {}
            });
        } else {
            console.error("Output file missing.");
            res.status(500).send('PDF file not found.');
        }
    });
});

// 9️⃣ PROTECT PDF
app.post('/protect-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    if (!req.body.password) return res.status(400).send('Password is required.');

    const inputPath = path.resolve(req.file.path);
    const password = req.body.password.replace(/"/g, '\\"'); 
    
    const outputFilename = `protected-${Date.now()}.pdf`;
    const outputPath = path.join(outputDir, outputFilename);

    console.log(`🔒 Encrypting with QPDF: ${req.file.originalname}`);

    const command = `${QPDF_PATH} --encrypt "${password}" "${password}" 256 -- "${inputPath}" "${outputPath}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.unlinkSync(inputPath); } catch(e) {}

        if (error) {
            console.error("QPDF Error:", stderr || stdout);
            return res.status(500).send('Encryption failed. Is the PDF already encrypted?');
        }

        if (fs.existsSync(outputPath)) {
            console.log("✅ Success! PDF Secured.");
            res.download(outputPath, (err) => {
                try { fs.unlinkSync(outputPath); } catch(e) {}
            });
        } else {
            res.status(500).send('Output file was not created.');
        }
    });
});

// 🔟 UNLOCK PDF
app.post('/unlock-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    if (!req.body.password) return res.status(400).send('Password is required to unlock.');

    const inputPath = path.resolve(req.file.path);
    const password = req.body.password.replace(/"/g, '\\"'); 
    
    const outputFilename = `unlocked-${Date.now()}.pdf`;
    const outputPath = path.join(outputDir, outputFilename);

    console.log(`🔓 Attempting to Unlock: ${req.file.originalname}`);

    const command = `${QPDF_PATH} --password="${password}" --decrypt "${inputPath}" "${outputPath}"`;

    exec(command, (error, stdout, stderr) => {
        try { fs.unlinkSync(inputPath); } catch(e) {}

        if (error) {
            console.error("QPDF Unlock Error:", stderr || stdout);
            if ((stderr && stderr.includes('invalid password')) || error.code === 2) {
                return res.status(401).send('Incorrect password. Please try again.');
            }
            return res.status(500).send('Failed to unlock PDF.');
        }

        if (fs.existsSync(outputPath)) {
            console.log("✅ Success! PDF Unlocked.");
            res.download(outputPath, (err) => {
                try { fs.unlinkSync(outputPath); } catch(e) {}
            });
        } else {
            res.status(500).send('Output file was not created.');
        }
    });
});

// ==========================================
// 1️⃣1️⃣ WORD TO PDF (Fixed Extension & Command)
// ==========================================
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    const inputPath = path.resolve(req.file.path);
    const outputDirPath = path.resolve(outputDir);
    
    // 1. FIX: Add original extension (e.g., .docx) so LibreOffice recognizes it
    const originalExt = path.extname(req.file.originalname); 
    const inputPathWithExt = inputPath + originalExt;
    
    try {
        fs.renameSync(inputPath, inputPathWithExt);
    } catch (err) {
        console.error("Rename Error:", err);
        return res.status(500).send("File processing error.");
    }

    // Create Unique Profile
    const uniqueProfileDir = path.join(tempDir, `profile-${Date.now()}`);
    const profileUrl = 'file:///' + uniqueProfileDir.replace(/\\/g, '/');

    // 2. FIX: Simplified Command (Removed complex FilterData to prevent crashes)
    const command = `${SOFFICE_PATH} -env:UserInstallation="${profileUrl}" --headless --convert-to pdf --outdir "${outputDirPath}" "${inputPathWithExt}"`;

    console.log("Optimizing and Converting Word to PDF..."); 

    exec(command, (error, stdout, stderr) => {
        // Cleanup temp profile
        try { fs.rmSync(uniqueProfileDir, { recursive: true, force: true }); } catch(e) {}

        if (error) {
            console.error(`Conversion Error: ${error.message}`);
            // Cleanup input file on fail
            try { fs.unlinkSync(inputPathWithExt); } catch(e) {} 
            return res.status(500).send('Conversion failed. Check server console.');
        }

        // Generate expected output path
        // LibreOffice uses the filename from inputPathWithExt for the output
        const outputName = path.parse(inputPathWithExt).name + '.pdf';
        const generatedFile = path.join(outputDir, outputName);

        if (fs.existsSync(generatedFile)) {
            res.download(generatedFile, (err) => {
                // Cleanup files after download
                try {
                    fs.unlinkSync(inputPathWithExt); 
                    fs.unlinkSync(generatedFile); 
                } catch(e) {}
            });
        } else {
            console.error("Output not found. LibreOffice Log:", stdout);
            try { fs.unlinkSync(inputPathWithExt); } catch(e) {} 
            res.status(500).send('PDF was not created. Ensure fonts are installed.');
        }
    });
});

// ==========================================
// 🚀 START SERVER
// ==========================================
app.listen(port, () => {
    console.log(`🚀 Server running at http://localhost:${port}`);
    console.log(`👉 Test it here: http://localhost:${port}`);
});