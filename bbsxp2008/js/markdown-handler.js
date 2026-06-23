/**
 * BBSXP Markdown Handler
 * Provides Markdown parsing and rendering with XSS protection
 */

(function() {
    'use strict';

    // Configure Marked.js options
    if (typeof marked !== 'undefined') {
        marked.setOptions({
            breaks: true,        // Convert \n to <br>
            gfm: true,          // GitHub Flavored Markdown
            headerIds: false,   // Don't add IDs to headers
            mangle: false,      // Don't mangle email addresses
            sanitize: false     // We'll use DOMPurify instead
        });
    }

    /**
     * Parse Markdown to HTML with XSS protection
     * @param {string} markdown - Raw Markdown text
     * @returns {string} - Sanitized HTML
     */
    window.parseMarkdown = function(markdown) {
        if (!markdown) return '';

        // Check if libraries are loaded
        if (typeof marked === 'undefined') {
            console.error('Marked.js not loaded');
            return escapeHtml(markdown);
        }

        if (typeof DOMPurify === 'undefined') {
            console.error('DOMPurify not loaded');
            return escapeHtml(markdown);
        }

        try {
            // Parse Markdown to HTML
            var html = marked.parse(markdown);

            // Sanitize HTML to prevent XSS
            var clean = DOMPurify.sanitize(html, {
                ALLOWED_TAGS: [
                    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                    'p', 'br', 'hr',
                    'strong', 'b', 'em', 'i', 'u', 's', 'del',
                    'code', 'pre',
                    'blockquote',
                    'ul', 'ol', 'li',
                    'a', 'img',
                    'table', 'thead', 'tbody', 'tr', 'th', 'td'
                ],
                ALLOWED_ATTR: [
                    'href', 'src', 'alt', 'title', 'class'
                ],
                ALLOWED_URI_REGEXP: /^(?:(?:(?:f|ht)tps?|mailto|tel|callto|cid|xmpp):|[^a-z]|[a-z+.\-]+(?:[^a-z+.\-:]|$))/i
            });

            return clean;
        } catch (e) {
            console.error('Markdown parsing error:', e);
            return escapeHtml(markdown);
        }
    };

    /**
     * Render Markdown content in DOM elements
     * @param {string} selector - CSS selector for elements containing Markdown
     */
    window.renderMarkdown = function(selector) {
        var elements = document.querySelectorAll(selector || '.markdown-content');

        for (var i = 0; i < elements.length; i++) {
            var elem = elements[i];
            var markdown = elem.textContent || elem.innerText;
            var html = parseMarkdown(markdown);
            elem.innerHTML = html;
            elem.classList.add('markdown-rendered');
        }
    };

    /**
     * Initialize Markdown editor with preview
     * @param {string} textareaId - ID of textarea element
     * @param {string} previewId - ID of preview element
     */
    window.initMarkdownEditor = function(textareaId, previewId) {
        var textarea = document.getElementById(textareaId);
        var preview = document.getElementById(previewId);

        if (!textarea) return;

        // Update preview on input
        if (preview) {
            textarea.addEventListener('input', function() {
                var html = parseMarkdown(textarea.value);
                preview.innerHTML = html;
            });

            // Initial render
            var html = parseMarkdown(textarea.value);
            preview.innerHTML = html;
        }

        // Add toolbar buttons
        addMarkdownToolbar(textarea);
    };

    /**
     * Add formatting toolbar above textarea
     * @param {HTMLElement} textarea
     */
    function addMarkdownToolbar(textarea) {
        var toolbar = document.createElement('div');
        toolbar.className = 'markdown-toolbar';
        toolbar.innerHTML = [
            '<button type="button" onclick="insertMarkdown(\'**\', \'**\')" title="粗体">B</button>',
            '<button type="button" onclick="insertMarkdown(\'*\', \'*\')" title="斜体">I</button>',
            '<button type="button" onclick="insertMarkdown(\'~~\', \'~~\')" title="删除线">S</button>',
            '<button type="button" onclick="insertMarkdown(\'`\', \'`\')" title="代码">Code</button>',
            '<button type="button" onclick="insertMarkdown(\'[链接]()\', \'\')" title="链接">Link</button>',
            '<button type="button" onclick="insertMarkdown(\'![图片]()\', \'\')" title="图片">Img</button>',
            '<button type="button" onclick="insertMarkdown(\'\\n> \', \'\\n\')" title="引用">Quote</button>',
            '<button type="button" onclick="insertMarkdown(\'\\n- \', \'\\n\')" title="列表">List</button>',
            '<button type="button" onclick="showMarkdownHelp()" title="帮助">?</button>'
        ].join(' ');

        textarea.parentNode.insertBefore(toolbar, textarea);
    }

    /**
     * Insert Markdown syntax at cursor position
     * @param {string} before - Text before cursor
     * @param {string} after - Text after cursor
     */
    window.insertMarkdown = function(before, after) {
        var textarea = document.activeElement;
        if (textarea.tagName !== 'TEXTAREA') return;

        var start = textarea.selectionStart;
        var end = textarea.selectionEnd;
        var text = textarea.value;
        var selectedText = text.substring(start, end);

        var newText = text.substring(0, start) + before + selectedText + after + text.substring(end);
        textarea.value = newText;

        // Set cursor position
        var newPos = start + before.length + selectedText.length;
        textarea.setSelectionRange(newPos, newPos);
        textarea.focus();

        // Trigger input event for preview update
        var event = new Event('input', { bubbles: true });
        textarea.dispatchEvent(event);
    };

    /**
     * Show Markdown syntax help
     */
    window.showMarkdownHelp = function() {
        var help = [
            'Markdown 语法帮助：',
            '',
            '**粗体** 或 __粗体__',
            '*斜体* 或 _斜体_',
            '~~删除线~~',
            '`代码`',
            '',
            '# 一级标题',
            '## 二级标题',
            '### 三级标题',
            '',
            '[链接文字](网址)',
            '![图片说明](图片网址)',
            '',
            '- 无序列表',
            '1. 有序列表',
            '',
            '> 引用文字',
            '',
            '```',
            '代码块',
            '```'
        ].join('\n');

        alert(help);
    };

    /**
     * Escape HTML special characters
     * @param {string} text
     * @returns {string}
     */
    function escapeHtml(text) {
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Auto-render on page load
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            renderMarkdown('.markdown-content');
        });
    } else {
        renderMarkdown('.markdown-content');
    }

})();
