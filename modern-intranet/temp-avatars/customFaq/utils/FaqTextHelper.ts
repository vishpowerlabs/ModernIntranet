/**
 * Helper utilities for text processing inside the FAQ web part.
 * These utilities keep CustomFaqWebPart smaller and easier to maintain.
 */
export class FaqTextHelper {
  /**
   * Highlight query matches within text.
   */
  public static highlightText(text: string, query: string, highlightClass: string, isHtml: boolean): string {
    if (!text || !query) {
      return text;
    }

    const trimmedQuery = query.trim();
    if (trimmedQuery === '') {
      return text;
    }

    if (isHtml) {
      return this._highlightInHtml(text, trimmedQuery, highlightClass);
    }

    const regex = new RegExp('(' + this.escapeRegExp(trimmedQuery) + ')', 'gi');
    return text.replace(regex, '<mark class="' + highlightClass + '">$1</mark>');
  }

  /**
   * Strip HTML content down to plain text.
   */
  public static stripHtmlTags(input: string): string {
    if (!input) {
      return '';
    }

    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + input + '</div>', 'text/html');
      const container = doc.body.firstChild as HTMLElement;
      if (!container) {
        return this._fallbackStripTags(input);
      }

      const result = container.textContent || '';
      return this._decodeHtmlEntities(result);
    } catch (error) {
      console.warn('DOMParser failed, using fallback:', error);
      return this._fallbackStripTags(input);
    }
  }

  /**
   * Escape regex special characters.
   */
  public static escapeRegExp(text: string): string {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  private static _highlightInHtml(html: string, query: string, highlightClass: string): string {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + html + '</div>', 'text/html');
      const container = doc.body.firstChild as HTMLElement;
      if (!container) {
        return html;
      }

      const walker = doc.createTreeWalker(container, NodeFilter.SHOW_TEXT, null);
      const textNodes: Text[] = [];
      let node: Text | null;
      while ((node = walker.nextNode() as Text | null)) {
        textNodes.push(node);
      }

      const lowerQuery = query.toLowerCase();
      let i = 0;
      while (i < textNodes.length) {
        const textNode = textNodes[i];
        const textContent = textNode.textContent || '';
        const lowerText = textContent.toLowerCase();
        const index = lowerText.indexOf(lowerQuery);

        if (index !== -1 && textNode.parentNode) {
          const before = textContent.substring(0, index);
          const match = textContent.substring(index, index + query.length);
          const after = textContent.substring(index + query.length);

          const fragment = doc.createDocumentFragment();
          if (before) {
            fragment.appendChild(doc.createTextNode(before));
          }

          const mark = doc.createElement('mark');
          mark.className = highlightClass;
          mark.textContent = match;
          fragment.appendChild(mark);

          if (after) {
            fragment.appendChild(doc.createTextNode(after));
          }

          textNode.parentNode.replaceChild(fragment, textNode);
        }
        i++;
      }

      return container.innerHTML;
    } catch (error) {
      console.warn('HTML highlighting failed:', error);
      return html;
    }
  }

  private static _fallbackStripTags(input: string): string {
    let result = '';
    let inTag = false;
    const chars = input.split('');
    let idx = 0;
    const len = chars.length;

    while (idx < len) {
      const char = chars[idx];
      if (char === '<') {
        inTag = true;
      } else if (char === '>') {
        inTag = false;
      } else if (!inTag) {
        result += char;
      }
      idx++;
    }

    return this._decodeHtmlEntities(result);
  }

  private static _decodeHtmlEntities(text: string): string {
    let result = text;
    result = result.split('&nbsp;').join(' ');
    result = result.split('&amp;').join('&');
    result = result.split('&lt;').join('<');
    result = result.split('&gt;').join('>');
    result = result.split('&quot;').join('"');
    result = result.split('&#39;').join("'");
    result = result.split('&#x27;').join("'");
    result = result.split('&apos;').join("'");
    return result;
  }
}
