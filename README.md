# Tunisian Ministry of Finance Document Extractor

A production-ready Flask web application that extracts structured data from Tunisian Ministry of Finance documents using OpenAI Vision API and generates XLSX files.

## Features

- 📤 Upload multiple photos (JPG/PNG) of Ministry documents
- 🤖 Automatic data extraction using OpenAI Vision API
- 📊 Generate structured XLSX files per document
- 📥 Download individual or batch XLSX files
- 🔍 Support for two document types:
  - **ACTES_SOCIETES** (Company Acts)
  - **BIENS_IMMOBILIERS_ACHETEUR** (Real Estate - Buyer Quality)
- ✨ Automatic detection and separation of multiple document types in a single image

## Document Types Supported

1. **"RECOUPEMENTS SUR ACTES D'ENREGISTREMENTS"**
2. **"Actes sur les Sociétés – Relatifs à toutes les années"**
3. **"BIENS IMMOBILIERS (Qualité Acheteur)"**

## Installation

### Prerequisites

- Python 3.8 or higher
- OpenAI API key

### Setup Steps

1. **Clone or download this repository**

2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**:
   - Copy `.env.example` to `.env`
   - Edit `.env` and add your OpenAI API key:
     ```
     OPEN_AI_API_KEY=sk-your-actual-api-key-here
     SECRET_KEY=your-secret-key-here
     ```

5. **Run the application**:
   ```bash
   python app.py
   ```

6. **Access the application**:
   - Open your browser and navigate to: `http://localhost:5000`

## Usage

1. **Upload Documents**:
   - Click the upload area or drag and drop image files
   - Select multiple JPG/PNG files (max 50MB per file)
   - Click "Process Documents"

2. **View Results**:
   - After processing, you'll see a results page with:
     - List of processed photos
     - Generated XLSX files for each photo
     - Download links for individual files

3. **Download Files**:
   - Click on any XLSX filename to download individually
   - Click "Download All as ZIP" to get all files in one archive

## File Naming Convention

- Single document type: `photo_name.xlsx`
- Multiple document types in one image:
  - `photo_name__ACTES_SOCIETES.xlsx`
  - `photo_name__BIENS_IMMO.xlsx`

## XLSX File Structure

Each XLSX file contains:

**Rows 1-5 (Metadata)**:
- photo_name
- doc_type
- processed_at
- raison_sociale
- matricule_fiscal

**Row 6+ (Data Table)**:
- Column headers based on document type
- Extracted data rows

### ACTES_SOCIETES Columns:
- annee, ref_enregistrement, date_enregistrement, type_acte, date_acte
- matricule_fiscal_societe, raison_sociale_societe, capital_societe, forme_juridique
- apport_numeraire, apport_nature, apport_fonds_commerce, apport_incorporation
- apport_creances, apport_autres, total_apports, total_annuel, total_general

### BIENS_IMMOBILIERS Columns:
- annee, ref_enregistrement, date_enregistrement, numero_quittance
- type_acte, nature_acte, date_acte, nbr_parts
- vendeur_matricule_fiscal, vendeur_cin, vendeur_nom
- numero_bien, nature_et_adresse_bien, recette_et_date_origine
- surface_bien, montant_vente_bien, total_annuel

## Project Structure

```
.
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── .env.example          # Environment variables template
├── .env                  # Your environment variables (create this)
├── README.md             # This file
├── templates/
│   ├── index.html        # Upload form
│   └── results.html      # Results page
├── uploads/              # Temporary upload folder (auto-created)
└── outputs/              # Generated XLSX files (auto-created)
```

## API Endpoints

- `GET /` - Upload form
- `POST /upload` - Process uploaded files
- `GET /download/<batch_id>/<filename>` - Download individual XLSX
- `GET /download_all/<batch_id>` - Download all files as ZIP

## Error Handling

- File type validation (JPG/PNG only)
- File size limits (50MB max)
- Graceful OpenAI API error handling
- Mock fallback XLSX generation if extraction fails

## Production Deployment

For production deployment:

1. **Change the secret key** in `.env`
2. **Set `debug=False`** in `app.py`
3. **Use a production WSGI server** (e.g., Gunicorn):
   ```bash
   pip install gunicorn
   gunicorn -w 4 -b 0.0.0.0:5000 app:app
   ```
4. **Configure reverse proxy** (nginx/Apache)
5. **Set up SSL/HTTPS**
6. **Configure proper file storage** (consider cloud storage for uploads/outputs)

## Troubleshooting

### OpenAI API Errors
- Verify your API key is correct in `.env`
- Check your OpenAI account has sufficient credits
- Ensure the API key has access to GPT-4 Vision models

### File Upload Issues
- Check file size (max 50MB)
- Verify file format (JPG/PNG only)
- Ensure `uploads/` folder has write permissions

### XLSX Generation Issues
- Verify `openpyxl` is installed correctly
- Check `outputs/` folder has write permissions

## License

This project is provided as-is for internal use.

## Support

For issues or questions, please contact the development team.
