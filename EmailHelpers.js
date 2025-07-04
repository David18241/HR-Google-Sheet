/**
 * @fileoverview Helper functions for creating and sending emails.
 */

/**
 * Converts a Google Docs body to HTML format for email use.
 * @param {Body} body The Google Docs body element to convert
 * @returns {string} HTML representation of the document body
 */
function convertBodyToHtml(body) {
  var html = '';
  var lastProcessedText = new Set(); // Track processed items by their text content

  for (var i = 0; i < body.getNumChildren(); i++) {
    var element = body.getChild(i);

    switch(element.getType()) {
      case DocumentApp.ElementType.PARAGRAPH:
        var headingTag = getHeadingTag(element);
        if (headingTag) {
          html += `\n<div class="section" style="margin-left: 0 !important;">
            <${headingTag} style="${getParagraphStyle(element)}">${convertParagraphToHtml(element)}</${headingTag}>
          </div>\n`;
        } else {
          html += `<p style="${getParagraphStyle(element)}">${convertParagraphToHtml(element)}</p>`;
        }
        break;

      case DocumentApp.ElementType.LIST_ITEM:
        var listItem = element.asListItem();
        var itemText = listItem.getText().trim();

        // Skip if we've already processed this exact text recently
        if (!lastProcessedText.has(itemText)) {
          html += processListItem(listItem);
          lastProcessedText.add(itemText);

          // Keep set size manageable by removing older items
          if (lastProcessedText.size > 100) {
            lastProcessedText.clear();
          }
        }
        break;

      case DocumentApp.ElementType.TABLE:
        html += convertTableToHtml(element);
        break;
        // Add other element types as needed
      default:
         Logger.log(`Unsupported element type: ${element.getType()}`);
    }
  }
  return html;
}

function processListItem(listItem) {
  var html = '';
  var nesting = listItem.getNestingLevel();
  var type = listItem.getGlyphType();
  var listTag = type === DocumentApp.GlyphType.BULLET || type === DocumentApp.GlyphType.HOLLOW_BULLET || type === DocumentApp.GlyphType.SQUARE_BULLET ? 'ul' : 'ol';

  var prevSibling = listItem.getPreviousSibling();
  var isPrevListItem = prevSibling && prevSibling.getType() === DocumentApp.ElementType.LIST_ITEM;
  var prevNesting = isPrevListItem ? prevSibling.asListItem().getNestingLevel() : -1;
  var prevType = isPrevListItem ? prevSibling.asListItem().getGlyphType() : null;
  var prevListTag = prevType === DocumentApp.GlyphType.BULLET || prevType === DocumentApp.GlyphType.HOLLOW_BULLET || prevType === DocumentApp.GlyphType.SQUARE_BULLET ? 'ul' : 'ol';

  // Start new list or adjust nesting/type as needed
  if (!isPrevListItem || prevNesting < nesting || (prevNesting === nesting && listTag !== prevListTag)) {
      // Close previous list of same level if type changed
      if (isPrevListItem && prevNesting === nesting && listTag !== prevListTag) {
          html += `</${prevListTag}>`;
      }
    html += `<${listTag} style="margin-left: ${nesting * 20}px; list-style-type: ${getGlyphStyle(type)};">`;
  }

  // Add the list item itself
  html += `<li style="margin-bottom: 0.5em;">${convertParagraphToHtml(listItem)}</li>`;

  // Handle list ending or change in nesting/type
  var nextSibling = listItem.getNextSibling();
  var isNextListItem = nextSibling && nextSibling.getType() === DocumentApp.ElementType.LIST_ITEM;
  var nextNesting = isNextListItem ? nextSibling.asListItem().getNestingLevel() : -1;
   var nextType = isNextListItem ? nextSibling.asListItem().getGlyphType() : null;
   var nextListTag = nextType === DocumentApp.GlyphType.BULLET || nextType === DocumentApp.GlyphType.HOLLOW_BULLET || nextType === DocumentApp.GlyphType.SQUARE_BULLET ? 'ul' : 'ol';

  if (!isNextListItem || nextNesting < nesting || (nextNesting === nesting && listTag !== nextListTag)) {
    html += `</${listTag}>`;
  }

   // Close outer lists if nesting decreases significantly
   if (isNextListItem && nextNesting < nesting -1) {
       for (let i = nesting -1; i > nextNesting; i--) {
            html += `</${listTag}>`;
       }
   }


  return html;
}

/**
 * Gets the CSS list-style-type for a Google Docs glyph type.
 * @param {DocumentApp.GlyphType} glyphType The glyph type from Google Docs
 * @returns {string} CSS list-style-type value
 */
function getGlyphStyle(glyphType) {
    switch (glyphType) {
        case DocumentApp.GlyphType.BULLET: return 'disc';
        case DocumentApp.GlyphType.HOLLOW_BULLET: return 'circle';
        case DocumentApp.GlyphType.SQUARE_BULLET: return 'square';
        case DocumentApp.GlyphType.NUMBER: return 'decimal';
        case DocumentApp.GlyphType.LATIN_LOWER: return 'lower-latin';
        case DocumentApp.GlyphType.LATIN_UPPER: return 'upper-latin';
        case DocumentApp.GlyphType.ROMAN_LOWER: return 'lower-roman';
        case DocumentApp.GlyphType.ROMAN_UPPER: return 'upper-roman';
        default: return 'disc';
    }
}

function getHeadingTag(paragraph) {
  var headingLevel = paragraph.getHeading();
  switch(headingLevel) {
    case DocumentApp.ParagraphHeading.HEADING1: return 'h1';
    case DocumentApp.ParagraphHeading.HEADING2: return 'h2';
    case DocumentApp.ParagraphHeading.HEADING3: return 'h3';
    case DocumentApp.ParagraphHeading.HEADING4: return 'h4';
    case DocumentApp.ParagraphHeading.HEADING5: return 'h5';
    case DocumentApp.ParagraphHeading.HEADING6: return 'h6';
    default: return null;
  }
}

function getParagraphStyle(paragraph) {
  var styles = [];

  // Basic margin/padding
  styles.push('margin: 0.5em 0');
  styles.push('padding: 0');

  // Alignment
  var alignment = paragraph.getAlignment();
   if (alignment) {
       switch(alignment) {
           case DocumentApp.HorizontalAlignment.CENTER: styles.push('text-align: center'); break;
           case DocumentApp.HorizontalAlignment.RIGHT: styles.push('text-align: right'); break;
           case DocumentApp.HorizontalAlignment.JUSTIFY: styles.push('text-align: justify'); break;
           case DocumentApp.HorizontalAlignment.LEFT:
              break;
       }
   }

  // Heading specific styles
  if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
    styles.push('font-weight: bold');
  }

  // Line spacing
  var spacing = paragraph.getLineSpacing();
  if (spacing) {
    styles.push('line-height: ' + spacing);
  }

  // Indentation
  var indentStart = paragraph.getIndentStart();
  var indentEnd = paragraph.getIndentEnd();
  var indentFirst = paragraph.getIndentFirstLine();

  if (indentStart) styles.push('padding-left: ' + indentStart + 'pt');
  if (indentEnd) styles.push('padding-right: ' + indentEnd + 'pt');
  if (indentFirst) styles.push('text-indent: ' + indentFirst + 'pt');


  return styles.join('; ');
}

function convertParagraphToHtml(paragraph) {
  var html = '';
  var text = paragraph.getText();
  // Handle empty paragraphs
  if (text === '') {
      return ' '; // Return non-breaking space for empty paragraphs to maintain spacing
  }
  var numChildren = paragraph.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = paragraph.getChild(i);

    if (child.getType() == DocumentApp.ElementType.TEXT) {
      html += convertTextToHtml(child.asText());
    } else if (child.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
        // Basic handling for images - might need refinement
        Logger.log("Inline image found - basic HTML conversion.");
        html += '[Image]'; // Or try to get URL/alt text if possible, complex
    }
    // Add other element types if needed (e.g., footnotes, equations)
  }
  return html;
}


/**
 * Converts a text element to HTML with proper formatting.
 * @param {Text} textElement The text element to convert
 * @returns {string} HTML representation of the text
 */
function convertTextToHtml(textElement) {
  var html = '';
  var text = textElement.getText();
  if (!text) return ''; // Handle empty text elements

  var attributes = textElement.getTextAttributeIndices();
  var lastIndex = 0;

  for (var i = 0; i < attributes.length; i++) {
    var index = attributes[i];
    // Process the segment before this attribute change
    html += processTextSegment(textElement, lastIndex, index);
    lastIndex = index;
  }
  // Process the final segment after the last attribute change
  html += processTextSegment(textElement, lastIndex, text.length);

  return html;
}

/**
 * Processes a segment of text with consistent formatting.
 * @param {Text} textElement The text element
 * @param {number} start Start index
 * @param {number} end End index
 * @returns {string} HTML for this text segment
 */
function processTextSegment(textElement, start, end) {
    if (start >= end) return ''; // Nothing to process

    var segmentText = textElement.getText().substring(start, end);
    var style = textElement.getAttributes(start); // Get attributes at the start of the segment

    var html = escapeHtml(segmentText);

    // Apply styling tags - order matters for proper nesting
    if (style[DocumentApp.Attribute.LINK_URL]) {
        html = '<a href="' + style[DocumentApp.Attribute.LINK_URL] + '">' + html + '</a>';
    }
    if (style[DocumentApp.Attribute.UNDERLINE]) {
        html = '<u>' + html + '</u>';
    }
     if (style[DocumentApp.Attribute.STRIKETHROUGH]) {
        html = '<s>' + html + '</s>';
    }
    if (style[DocumentApp.Attribute.ITALIC]) {
        html = '<em>' + html + '</em>';
    }
    if (style[DocumentApp.Attribute.BOLD]) {
        html = '<strong>' + html + '</strong>';
    }


    // Apply span for color, background, font size, font family
    var spanStyles = [];
    if (style[DocumentApp.Attribute.FOREGROUND_COLOR] && style[DocumentApp.Attribute.FOREGROUND_COLOR] !== '#000000') {
        spanStyles.push('color: ' + style[DocumentApp.Attribute.FOREGROUND_COLOR]);
    }
     if (style[DocumentApp.Attribute.BACKGROUND_COLOR] && style[DocumentApp.Attribute.BACKGROUND_COLOR] !== '#ffffff') {
        spanStyles.push('background-color: ' + style[DocumentApp.Attribute.BACKGROUND_COLOR]);
    }
    if (style[DocumentApp.Attribute.FONT_SIZE]) {
        spanStyles.push('font-size: ' + style[DocumentApp.Attribute.FONT_SIZE] + 'pt');
    }
     if (style[DocumentApp.Attribute.FONT_FAMILY] && style[DocumentApp.Attribute.FONT_FAMILY] !== 'Arial') {
        spanStyles.push('font-family: "' + style[DocumentApp.Attribute.FONT_FAMILY] + '"');
    }
    // Add vertical offset for sub/superscript if needed
    if (DocumentApp.VerticalTextAlignment && style[DocumentApp.Attribute.VERTICAL_OFFSET]) {
      try {
        if (style[DocumentApp.Attribute.VERTICAL_OFFSET] === DocumentApp.VerticalTextAlignment.SUPERSCRIPT) {
          html = '<sup>' + html + '</sup>';
        } else if (style[DocumentApp.Attribute.VERTICAL_OFFSET] === DocumentApp.VerticalTextAlignment.SUBSCRIPT) {
          html = '<sub>' + html + '</sub>';
        }
      } catch (error) {
        // Silently handle any vertical alignment issues without logging warnings
        // The original text formatting will be preserved
      }
    }


    if (spanStyles.length > 0) {
        html = '<span style="' + spanStyles.join('; ') + '">' + html + '</span>';
    }

    return html;
}


/**
 * Legacy function for tag mapping (kept for compatibility).
 * @param {string} style Style name
 * @returns {string} HTML tag
 */
function getTag(style) {
  switch(style) {
    case 'bold': return 'strong';
    case 'italic': return 'em';
    case 'underline': return 'u';
    default: return '';
  }
}

function escapeHtml(unsafe) {
  if (typeof unsafe !== 'string') {
      Logger.log("Warning: escapeHtml called with non-string value: " + typeof unsafe);
      return unsafe;
  }
  return unsafe
    .replace(/&/g, "&")
    .replace(/</g, "<")
    .replace(/>/g, ">")
    .replace(/"/g, "&quot")
    .replace(/'/g, "'");
}

/**
 * Converts a table element to HTML format.
 * @param {Table} tableElement The table element to convert
 * @returns {string} HTML representation of the table
 */
function convertTableToHtml(tableElement) {
    Logger.log("Table found - converting to HTML.");
    var html = '<table border="1" style="border-collapse: collapse; width: 100%;">';
    var numRows = tableElement.getNumRows();
    for (var i = 0; i < numRows; i++) {
        html += '<tr>';
        var row = tableElement.getRow(i);
        var numCells = row.getNumCells();
        for (var j = 0; j < numCells; j++) {
            var cell = row.getCell(j);
            var cellTag = (i === 0) ? 'th' : 'td';
            var cellContent = '';
             for (var k = 0; k < cell.getNumChildren(); k++) {
                var element = cell.getChild(k);
                 switch(element.getType()) {
                    case DocumentApp.ElementType.PARAGRAPH:
                         cellContent += `<div style="${getParagraphStyle(element)}">${convertParagraphToHtml(element)}</div>`;
                         break;
                    case DocumentApp.ElementType.LIST_ITEM:
                        cellContent += processListItem(element.asListItem());
                        break;
                     case DocumentApp.ElementType.TABLE:
                         cellContent += convertTableToHtml(element.asTable());
                         break;
                    default:
                        cellContent += element.getText();
                 }
            }
             var cellStyle = '';
             html += `<${cellTag} style="${cellStyle}">${cellContent}</${cellTag}>`;
        }
        html += '</tr>';
    }
    html += '</table>';
    return html;
}


// --- Email Template Functions ---


/**
 * Creates an email draft or sends an email based on a Google Doc template.
 * Handles placeholder replacement and Doc-to-HTML conversion.
 *
 * @param {string} templateId The ID of the Google Doc template.
 * @param {string} recipient The email address of the recipient.
 * @param {string} subject The email subject line.
 * @param {Object} placeholderValues An object where keys are placeholders (e.g., "{{FirstName}}")
 *                                   and values are the text replacements. For links, the value should be
 *                                   an object: { text: "Link Display Text", url: "http://..." }.
 * @param {boolean} [isDraft=true] If true, creates a draft; if false, sends the email directly.
 * @param {string} [tempCopyName="Temp Email Copy"] Optional name for the temporary doc copy.
 * @returns {boolean} True if the operation was successful, false otherwise.
 */
function createEmailFromTemplate(templateId, recipient, subject, placeholderValues, isDraft = true, tempCopyName = "Temp Email Copy") {
  let copyDocId = null; // Keep track of the copy ID for cleanup
  try {
    // 1. Make a copy of the template
    const templateDoc = DriveApp.getFileById(templateId);
    const copyDoc = templateDoc.makeCopy(`${tempCopyName} - ${new Date().toISOString()}`); // Add timestamp for uniqueness
    copyDocId = copyDoc.getId();

    // 2. Open the copy and replace placeholders
    const doc = DocumentApp.openById(copyDocId);
    const body = doc.getBody();

    for (const placeholder in placeholderValues) {
      if (placeholderValues.hasOwnProperty(placeholder)) {
        const value = placeholderValues[placeholder];
        if (typeof value === 'object' && value !== null && value.hasOwnProperty('text') && value.hasOwnProperty('url')) {
          // Handle link replacement
          replacePlaceholderWithLink(body, placeholder, value.text, value.url);
        } else {
          // Handle simple text replacement
          body.replaceText(placeholder, value || ''); // Replace null/undefined with empty string
        }
      }
    }
    doc.saveAndClose();

    // 3. Convert the Doc body to HTML
    // Re-open the doc to get the final body content after replacements and saves.
    const finalDoc = DocumentApp.openById(copyDocId);
    const htmlBody = convertBodyToHtml(finalDoc.getBody());
     if (!htmlBody) {
        Logger.log(`Warning: Generated HTML body for template ${templateId} is empty or null.`);
        // Optionally decide whether to proceed or return false
    }

    // 4. Create Draft or Send Email
    const options = { htmlBody: htmlBody };
    if (isDraft) {
      GmailApp.createDraft(recipient, subject, '', options);
      Logger.log(`Created draft email "${subject}" for ${recipient} using template ${templateId}`);
    } else {
      MailApp.sendEmail(recipient, subject, '', options);
      Logger.log(`Sent email "${subject}" to ${recipient} using template ${templateId}`);
    }

    // 5. Delete the temporary copy
    DriveApp.getFileById(copyDocId).setTrashed(true);
    copyDocId = null; // Nullify ID after successful trashing
    return true;

  } catch (error) {
    Logger.log(`Error creating/sending email "${subject}" for ${recipient} from template ${templateId}: ${error.message} \nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Failed to ${isDraft ? 'create draft' : 'send email'} "${subject}". Check logs.`);

    // Attempt to clean up the temporary file if it was created
    if (copyDocId) {
      try {
        DriveApp.getFileById(copyDocId).setTrashed(true);
        Logger.log(`Cleaned up temporary file: ${copyDocId}`);
      } catch (cleanupError) {
        Logger.log(`Error cleaning up temporary file ${copyDocId}: ${cleanupError.message}`);
      }
    }
    return false;
  }
}