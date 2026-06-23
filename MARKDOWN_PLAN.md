# Markdown Support Implementation Plan

## Overview
Add Markdown formatting support to BBSXP forum posts with client-side rendering.

## Components Needed

### 1. Markdown Parser Library
**Recommended**: Marked.js (https://marked.js.org/)
- Lightweight (~20KB minified)
- Fast parsing
- Security features (DOMPurify integration)
- Wide browser support

### 2. Implementation Files

#### A. JavaScript Library Files (bbsxp2008/js/)
- `marked.min.js` - Markdown parser
- `dompurify.min.js` - XSS sanitizer
- `markdown-handler.js` - Custom integration

#### B. CSS Styling (bbsxp2008/css/)
- `markdown-content.css` - Markdown rendered styles

#### C. Modified ASP Files
1. **AddPost.asp / AddTopic.asp** - Add Markdown editor
2. **EditPost.asp** - Edit with Markdown
3. **ShowPost.asp** - Render Markdown to HTML
4. **BBSXP_Class.asp** - Add Markdown processing function

### 3. Database Changes
**Option 1**: Add new field `PostFormat` (0=HTML, 1=Markdown)
**Option 2**: Auto-detect format by content

### 4. Security Considerations
- **Input**: Store raw Markdown in database (with HTMLEncode)
- **Output**: Parse Markdown → Sanitize with DOMPurify → Display
- **XSS Prevention**: DOMPurify removes dangerous HTML
- **SQL Injection**: Already protected by existing SqlString()

## Implementation Steps

### Phase 1: Add Libraries
1. Download marked.min.js
2. Download dompurify.min.js  
3. Create markdown-handler.js
4. Create markdown-content.css

### Phase 2: Backend Support
1. Add `MarkdownToHtml()` function in BBSXP_Class.asp
2. Modify content encoding in AddPost.asp
3. Update ShowPost.asp to render Markdown

### Phase 3: Frontend Editor
1. Add Markdown preview pane
2. Add formatting toolbar (bold, italic, link, etc.)
3. Add help/cheatsheet popup

### Phase 4: Testing
1. Test XSS protection
2. Test SQL injection protection
3. Test browser compatibility
4. Test existing posts compatibility

## Markdown Syntax Support

### Basic Formatting
- **Bold**: `**text**` or `__text__`
- *Italic*: `*text*` or `_text_`
- ~~Strikethrough~~: `~~text~~`
- `Code`: `` `code` ``

### Headings
```
# H1
## H2
### H3
```

### Lists
```
- Bullet item
1. Numbered item
```

### Links & Images
```
[Link text](url)
![Image alt](image-url)
```

### Code Blocks
```
\`\`\`javascript
code here
\`\`\`
```

### Quotes
```
> Quote text
```

### Tables (Optional)
```
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
```

## Security Model

```
User Input (Markdown)
    ↓
HTMLEncode() → Store in DB
    ↓
Retrieve from DB
    ↓
Marked.js Parse → HTML
    ↓
DOMPurify Sanitize → Safe HTML
    ↓
Display to User
```

## Backward Compatibility

### For Existing Posts
- Detect if post contains Markdown syntax
- If not, treat as plain text/HTML
- Add migration script (optional)

### Database Schema
```sql
-- Option: Add format field
ALTER TABLE BBSXP_Posts ADD PostFormat TINYINT DEFAULT 0
-- 0 = Plain/HTML, 1 = Markdown
```

## File Structure

```
bbsxp2008/
├── js/
│   ├── marked.min.js          (NEW)
│   ├── dompurify.min.js       (NEW)
│   └── markdown-handler.js    (NEW)
├── css/
│   └── markdown-content.css   (NEW)
├── AddPost.asp                (MODIFY)
├── AddTopic.asp               (MODIFY)
├── EditPost.asp               (MODIFY)
├── ShowPost.asp               (MODIFY)
└── BBSXP_Class.asp           (MODIFY)
```

## Next Steps

1. ✅ Create implementation plan
2. ⏳ Download and add JavaScript libraries
3. ⏳ Create markdown-handler.js
4. ⏳ Create CSS styling
5. ⏳ Modify backend ASP files
6. ⏳ Add frontend editor enhancements
7. ⏳ Security testing
8. ⏳ Documentation

## Estimated Timeline
- Library setup: 1 hour
- Backend integration: 2-3 hours
- Frontend editor: 2-3 hours
- Testing & debugging: 2 hours
- **Total**: ~8 hours

## Questions to Decide

1. Should Markdown be optional or mandatory?
2. Should we migrate existing posts?
3. What Markdown extensions to enable? (tables, task lists, etc.)
4. Should we add a live preview pane?

---

*Document Created: 2024*
*Status: Planning Phase*
