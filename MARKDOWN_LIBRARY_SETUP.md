# Markdown Support - Library Setup Guide

## Required Third-Party Libraries

### 1. Marked.js (Markdown Parser)

**Download**: https://cdn.jsdelivr.net/npm/marked/marked.min.js

**Manual Steps**:
1. Visit: https://github.com/markedjs/marked/releases
2. Download latest `marked.min.js`
3. Save to: `bbsxp2008/js/marked.min.js`

**CDN Alternative**:
```html
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
```

**Version**: Latest stable (recommend v4.0+)  
**License**: MIT  
**Size**: ~20KB minified

---

### 2. DOMPurify (XSS Sanitizer)

**Download**: https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js

**Manual Steps**:
1. Visit: https://github.com/cure53/DOMPurify/releases
2. Download latest `purify.min.js`
3. Save to: `bbsxp2008/js/dompurify.min.js`

**CDN Alternative**:
```html
<script src="https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js"></script>
```

**Version**: Latest stable (recommend v2.3+)  
**License**: Apache 2.0 or MPL 2.0  
**Size**: ~24KB minified

---

## Quick Download Commands

### Using curl
```bash
cd bbsxp2008/js/

# Download Marked.js
curl -o marked.min.js https://cdn.jsdelivr.net/npm/marked/marked.min.js

# Download DOMPurify
curl -o dompurify.min.js https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js
```

### Using wget
```bash
cd bbsxp2008/js/

# Download Marked.js
wget -O marked.min.js https://cdn.jsdelivr.net/npm/marked/marked.min.js

# Download DOMPurify
wget -O dompurify.min.js https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js
```

### Using PowerShell (Windows)
```powershell
cd bbsxp2008\js\

# Download Marked.js
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/marked/marked.min.js" -OutFile "marked.min.js"

# Download DOMPurify
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js" -OutFile "dompurify.min.js"
```

---

## Integration in ASP Pages

### Include in Header (Add to Setup.asp or BBSXP_Class.asp)

```html
<!-- Markdown Support Libraries -->
<script src="js/marked.min.js"></script>
<script src="js/dompurify.min.js"></script>
<script src="js/markdown-handler.js"></script>
<link rel="stylesheet" href="css/markdown-content.css">
```

---

## File Structure After Setup

```
bbsxp2008/
├── js/
│   ├── marked.min.js          ⬅️ Download this (20KB)
│   ├── dompurify.min.js       ⬅️ Download this (24KB)
│   └── markdown-handler.js    ✅ Already created
├── css/
│   └── markdown-content.css   ✅ Already created
└── [ASP files to modify]
```

---

## Verification

After downloading, verify files exist:

```bash
# Check if files downloaded successfully
ls -lh bbsxp2008/js/marked.min.js
ls -lh bbsxp2008/js/dompurify.min.js

# Verify file sizes (approximate)
# marked.min.js: ~20KB
# dompurify.min.js: ~24KB
```

---

## Next Steps

After downloading libraries:

1. ✅ Libraries downloaded
2. ⏳ Modify `ShowPost.asp` to render Markdown
3. ⏳ Modify `AddPost.asp` to add Markdown editor
4. ⏳ Modify `AddTopic.asp` for new topics
5. ⏳ Modify `EditPost.asp` for editing
6. ⏳ Test and verify

---

## Alternative: Use CDN (No Download Needed)

If you prefer using CDN instead of local files, simply include these URLs in your pages:

```html
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/dompurify/dist/purify.min.js"></script>
<script src="js/markdown-handler.js"></script>
<link rel="stylesheet" href="css/markdown-content.css">
```

**Pros**:
- No download needed
- Automatic caching
- CDN performance

**Cons**:
- Requires internet connection
- External dependency
- Privacy considerations

---

## License Information

### Marked.js
- **License**: MIT
- **Copyright**: Christopher Jeffrey
- **URL**: https://github.com/markedjs/marked

### DOMPurify
- **License**: Apache 2.0 or MPL 2.0
- **Copyright**: Cure53
- **URL**: https://github.com/cure53/DOMPurify

### Our Code (markdown-handler.js)
- **License**: Same as BBSXP project
- **Copyright**: BBSXP Contributors

---

## Security Notes

1. **Always use both libraries together**
   - Marked.js: Parses Markdown
   - DOMPurify: Sanitizes output (prevents XSS)

2. **Never skip DOMPurify**
   - Without it, XSS attacks are possible
   - DOMPurify removes dangerous HTML/JavaScript

3. **Keep libraries updated**
   - Check for security updates regularly
   - Update at least quarterly

---

## Support & Troubleshooting

### Library not loading?
- Check file paths are correct
- Verify files downloaded completely
- Check browser console for errors

### Markdown not rendering?
- Verify libraries loaded before markdown-handler.js
- Check console for JavaScript errors
- Ensure elements have correct CSS class

### XSS concerns?
- Always load DOMPurify
- Test with known XSS payloads
- Review sanitization settings

---

*Document Created: 2024*  
*Status: Ready for Library Download*
