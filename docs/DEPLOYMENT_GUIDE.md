# Deployment Guide: DOCX Anonymizer to Streamlit Cloud

This guide walks through deploying the DOCX Anonymizer app to Streamlit Cloud with GitHub integration.

## Current Status

Phase 2 (Deployment Preparation) is COMPLETE:
- Deployment directory created at `/tmp/docx-anonymizer-app/`
- All necessary files prepared
- Git repository initialized
- Initial commit created

## Next Steps

### Phase 3: Create GitHub Repository and Push Code

Since you're in a local environment, you'll need to manually create the GitHub repository and push the code.

#### Option A: Using GitHub CLI (`gh`)

If you have `gh` installed:

```bash
cd /tmp/docx-anonymizer-app
gh auth login  # If not already authenticated
gh repo create docx-anonymizer-app --public --source=. --description="Web app for anonymizing Word documents with Excel-based mappings and PDF conversion" --push
```

#### Option B: Manual GitHub Setup (Recommended)

1. **Go to GitHub** and create a new repository:
   - Navigate to https://github.com/new
   - Repository name: `docx-anonymizer-app`
   - Description: `Web app for anonymizing Word documents with Excel-based mappings and PDF conversion`
   - Visibility: **Public**
   - Do NOT initialize with README, .gitignore, or license (we already have these)

2. **Push your local code** to GitHub:

```bash
cd /tmp/docx-anonymizer-app
git remote add origin https://github.com/svan-b/docx-anonymizer-app.git
git branch -M main
git push -u origin main
```

### Phase 4: Deploy to Streamlit Cloud

1. **Go to Streamlit Cloud**: https://share.streamlit.io/

2. **Sign in** with your GitHub account (svan-b)

3. **Deploy new app**:
   - Click "New app" button
   - Repository: `svan-b/docx-anonymizer-app`
   - Branch: `main`
   - Main file path: `app.py`
   - Click "Deploy!"

4. **Streamlit will automatically**:
   - Install LibreOffice (via `packages.txt`)
   - Install Python dependencies (via `requirements.txt`)
   - Apply Streamlit configuration (via `.streamlit/config.toml`)
   - Deploy your app

5. **Wait for deployment** (2-5 minutes):
   - Streamlit will show build logs
   - Watch for any errors (especially LibreOffice installation)
   - App URL will be: `https://<app-name>.streamlit.app`

6. **Test the deployment**:
   - Upload `sample_requirements.xlsx`
   - Upload a test DOCX file
   - Verify anonymization works
   - Verify PDF conversion works
   - Test both checkboxes:
     - "Remove all images from document"
     - "Clear headers/footers (for presentations with logos)"

### Phase 5: Set Up Local-to-Cloud Sync Workflow

After deployment is complete, any changes you make locally can be pushed to Streamlit Cloud using this workflow:

```bash
# 1. Make changes to files in ui/ directory (your local environment)
cd /mnt/c/Users/vanbo/Development/Projects/xAI/anonymous/vdr-processor-docx/ui/

# 2. Copy updated files to deployment directory
cp docx_anonymizer_app.py /tmp/docx-anonymizer-app/app.py
cp process_adobe_word_files.py /tmp/docx-anonymizer-app/process_adobe_word_files.py

# 3. Commit and push to GitHub
cd /tmp/docx-anonymizer-app
git add .
git commit -m "Update: [description of changes]"
git push origin main

# 4. Streamlit Cloud will automatically redeploy (1-2 minutes)
```

### Creating a Sync Script

For convenience, create `/tmp/sync_to_streamlit.sh`:

```bash
#!/bin/bash

# Sync local changes to Streamlit Cloud deployment

echo "Syncing local ui/ files to deployment directory..."

# Copy files
cp /mnt/c/Users/vanbo/Development/Projects/xAI/anonymous/vdr-processor-docx/ui/docx_anonymizer_app.py /tmp/docx-anonymizer-app/app.py
cp /mnt/c/Users/vanbo/Development/Projects/xAI/anonymous/vdr-processor-docx/ui/process_adobe_word_files.py /tmp/docx-anonymizer-app/process_adobe_word_files.py

# Navigate to deployment directory
cd /tmp/docx-anonymizer-app

# Check for changes
if [[ -z $(git status -s) ]]; then
    echo "No changes to sync."
    exit 0
fi

# Show changes
echo "Changes detected:"
git status -s

# Commit and push
read -p "Enter commit message: " commit_msg
git add .
git commit -m "$commit_msg"
git push origin main

echo "✓ Changes pushed to GitHub. Streamlit Cloud will redeploy automatically."
```

Make it executable:
```bash
chmod +x /tmp/sync_to_streamlit.sh
```

## Troubleshooting

### LibreOffice Issues

If PDF conversion fails on Streamlit Cloud:

**Symptom**: Files process but PDF download is empty or conversion fails

**Solution**:
- Check build logs for LibreOffice installation errors
- Verify `packages.txt` contains `libreoffice`
- LibreOffice may not support very long filenames (>100 chars) or password-protected files
- Display error message to user if conversion fails

### File Upload Size Limits

Streamlit Cloud has upload limits:
- Default: 200MB (configured in `.streamlit/config.toml`)
- If you need larger files, contact Streamlit support

### Memory Issues

If app crashes with large files:
- Consider reducing batch size
- Process files one at a time instead of all at once
- Use Streamlit's progress indicators to show processing status

### Git Sync Issues

If automatic redeployment doesn't trigger:
- Check GitHub webhook settings in Streamlit Cloud
- Manually trigger redeploy from Streamlit Cloud dashboard
- Verify branch name is correct (main vs master)

## File Structure Summary

```
/tmp/docx-anonymizer-app/
├── .git/                          # Git repository
├── .gitignore                      # Git ignore patterns
├── .streamlit/
│   └── config.toml                 # Streamlit configuration
├── app.py                          # Main Streamlit app
├── process_adobe_word_files.py    # Backend processing
├── packages.txt                    # System dependencies (LibreOffice)
├── requirements.txt                # Python dependencies
├── sample_requirements.xlsx        # Example mapping file
├── README.md                       # Documentation
└── DEPLOYMENT_GUIDE.md            # This guide

```

## Post-Deployment Checklist

- [ ] GitHub repository created and code pushed
- [ ] Streamlit Cloud app deployed successfully
- [ ] LibreOffice installed and working
- [ ] Sample file tested successfully
- [ ] Both checkboxes functional
- [ ] PDF conversion working
- [ ] ZIP downloads working
- [ ] Sync script created and tested
- [ ] App URL shared/bookmarked

## Your App URL

After deployment, your app will be accessible at:
```
https://docx-anonymizer-app-[hash].streamlit.app
```

You can customize this URL in Streamlit Cloud settings.
