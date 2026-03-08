import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

import styles from './CustomFaq.module.scss';
import { SharePointService, IListInfo, IColumnInfo, IFaqItem } from './services/SharePointService';
import { FaqTextHelper } from './utils/FaqTextHelper';

export interface ICustomFaqWebPartProps {
  webpartTitle: string;
  webpartDescription: string;
  selectedList: string;
  titleColumn: string;
  descriptionColumn: string;
  categoryColumn: string;
  topFaqColumn: string;
  topFaqColor: string;
  allowMultipleExpanded: boolean;
  webpartTitleFontSize: string;
  webpartDescriptionFontSize: string;
  faqTitleFontSize: string;
  faqDescriptionFontSize: string;
  tabFontSize: string;
  enableSearch: boolean;
  searchInHeader: boolean;
  highlightSearchMatches: boolean;
  searchInAnswers: boolean;
  showResultsCount: boolean;
  searchPlaceholder: string;
  subCategoryColumn: string;
  filterCategory: string;
}


export default class CustomFaqWebPart extends BaseClientSideWebPart<ICustomFaqWebPartProps> {

  private _spService!: SharePointService;
  private _lists: IListInfo[] = [];
  private _columns: IColumnInfo[] = [];
  private _faqItems: IFaqItem[] = [];
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;
  private _selectedCategory: string = 'All';
  private _categories: string[] = [];
  private _searchQuery: string = '';
  private readonly _topFaqCategoryKey: string = '__TOP_FAQ__';
  private _pendingDeepLinkId: number | null = null;
  private _activeDeepLinkId: number | null = null;

  /**
   * Initialize the web part
   */
  protected onInit(): Promise<void> {
    return super.onInit()
      .then((): void => {
        // Initialize SharePoint service
        this._spService = new SharePointService(this.context as WebPartContext);

        // Try to consume the ThemeProvider service for section background support
        try {
          this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
          if (this._themeProvider) {
            this._themeVariant = this._themeProvider.tryGetTheme();
            this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
          }
        } catch (error) {
          console.log('ThemeProvider not available:', error);
        }

        if (this._themeVariant) {
          this._setCSSVariables(this._themeVariant);
        }

        if (typeof window !== 'undefined') {
          window.addEventListener('hashchange', this._handleHashChange);
          this._setPendingDeepLinkFromHash(window.location.hash);
        }
      })
      .then(() => {
        return this._loadLists();
      })
      .then(() => {
        if (this.properties.selectedList) {
          return this._loadColumns().then(() => this._loadFaqItems());
        }
      })
      .catch((error: Error): void => {
        console.error('Error during initialization:', error);
      });
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    if (this._themeVariant) {
      this._setCSSVariables(this._themeVariant);
    }
    this.render();
  }

  private _setCSSVariables(theme: IReadonlyTheme): void {
    if (!this.domElement) {
      return;
    }

    try {
      if (theme.semanticColors) {
        const semanticColors: { [key: string]: string } = theme.semanticColors as { [key: string]: string };
        const semanticKeys = Object.keys(semanticColors);
        let i = 0;
        while (i < semanticKeys.length) {
          const key = semanticKeys[i];
          const value = semanticColors[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
          i++;
        }
      }

      if (theme.palette) {
        const palette: { [key: string]: string } = theme.palette as { [key: string]: string };
        const paletteKeys = Object.keys(palette);
        let j = 0;
        while (j < paletteKeys.length) {
          const key = paletteKeys[j];
          const value = palette[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
          j++;
        }
      }
    } catch (error) {
      console.error('Error setting CSS variables:', error);
    }
  }

  private _extractCategories(): void {
    const categorySet: { [key: string]: boolean } = {};
    this._categories = [];
    const filterCategory = this.properties.filterCategory;

    let i = 0;
    while (i < this._faqItems.length) {
      const item = this._faqItems[i];
      // If filtering by a main category, only consider items in that category
      if (filterCategory && item.category !== filterCategory) {
        i++;
        continue;
      }

      // If filtering by main category, use subCategory for tabs. Otherwise use category.
      const categoryToUse = filterCategory ? item.subCategory : item.category;

      if (categoryToUse && !categorySet[categoryToUse]) {
        categorySet[categoryToUse] = true;
        this._categories.push(categoryToUse);
      }
      i++;
    }

    this._categories.sort();
  }

  private _isTopFaqEnabled(): boolean {
    return !!(this.properties.topFaqColumn && this.properties.topFaqColumn.trim() !== '');
  }

  private _hasTopFaqItems(): boolean {
    let i = 0;
    const filterCategory = this.properties.filterCategory;
    while (i < this._faqItems.length) {
      if (filterCategory && this._faqItems[i].category !== filterCategory) {
        i++;
        continue;
      }

      if (this._faqItems[i].isTopFaq) {
        return true;
      }
      i++;
    }
    return false;
  }

  private _ensureValidCategorySelection(): void {
    const validCategories: { [key: string]: boolean } = {
      'All': true
    };

    if (this._isTopFaqEnabled()) {
      validCategories[this._topFaqCategoryKey] = true;
    }

    let i = 0;
    while (i < this._categories.length) {
      validCategories[this._categories[i]] = true;
      i++;
    }

    if (!validCategories[this._selectedCategory]) {
      this._selectedCategory = 'All';
    }
  }

  private _sanitizeColorInput(color: string | undefined): string {
    if (!color) {
      return '';
    }

    const trimmed = color.trim();
    if (trimmed === '') {
      return '';
    }

    const hexRegex = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
    const rgbRegex = /^rgba?\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}(?:\s*,\s*(0|0?\.\d+|1(\.0)?))?\s*\)$/;
    const cssNameRegex = /^[a-zA-Z]+$/;

    if (hexRegex.test(trimmed) || rgbRegex.test(trimmed) || cssNameRegex.test(trimmed)) {
      return trimmed;
    }

    return '';
  }

  private _getTopFaqColor(fallbackColor: string): string {
    const sanitized = this._sanitizeColorInput(this.properties.topFaqColor);
    if (sanitized) {
      return sanitized;
    }
    return fallbackColor;
  }

  private _getTopFaqBadgePalette(color: string): { badgeBackground: string; badgeTextColor: string; badgeBorder: string; badgeShadow: string } {
    const badgeBackground = color;
    const badgeBorder = this._hexToRgba(color, 0.35) || 'rgba(0, 0, 0, 0.3)';
    const badgeShadow = this._hexToRgba(color, 0.45) || 'rgba(0, 0, 0, 0.25)';
    const badgeTextColor = this._getContrastingTextColor(color);

    return {
      badgeBackground,
      badgeTextColor,
      badgeBorder,
      badgeShadow
    };
  }

  private _hexToRgba(color: string, alpha: number): string | null {
    if (!color) {
      return null;
    }
    const trimmed = color.trim();
    if (!trimmed || trimmed.charAt(0) !== '#') {
      return null;
    }

    let hex = trimmed.substring(1);
    if (hex.length === 3) {
      hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
    }

    if (hex.length !== 6 || hex.match(/[^0-9a-fA-F]/)) {
      return null;
    }

    const r = Number.parseInt(hex.substring(0, 2), 16);
    const g = Number.parseInt(hex.substring(2, 4), 16);
    const b = Number.parseInt(hex.substring(4, 6), 16);
    const normalizedAlpha = Math.max(0, Math.min(1, alpha));

    return 'rgba(' + r + ', ' + g + ', ' + b + ', ' + normalizedAlpha + ')';
  }

  private _parseHexColor(color: string): { r: number; g: number; b: number } | null {
    if (!color || color.charAt(0) !== '#') {
      return null;
    }

    let hex = color.substring(1);
    if (hex.length === 3) {
      hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
    }

    if (hex.length !== 6 || hex.match(/[^0-9a-fA-F]/)) {
      return null;
    }

    return {
      r: Number.parseInt(hex.substring(0, 2), 16),
      g: Number.parseInt(hex.substring(2, 4), 16),
      b: Number.parseInt(hex.substring(4, 6), 16)
    };
  }

  private _relativeLuminance(r: number, g: number, b: number): number {
    const srgb = [r, g, b].map((value: number) => {
      const normalized = value / 255;
      return normalized <= 0.03928 ? normalized / 12.92 : Math.pow((normalized + 0.055) / 1.055, 2.4);
    });

    return 0.2126 * srgb[0] + 0.7152 * srgb[1] + 0.0722 * srgb[2];
  }

  private _getContrastingTextColor(color: string): string {
    const rgb = this._parseHexColor(color);
    if (!rgb) {
      return '#ffffff';
    }

    const luminance = this._relativeLuminance(rgb.r, rgb.g, rgb.b);
    return luminance > 0.6 ? '#201f1e' : '#ffffff';
  }

  /**
   * Get filtered FAQ items based on selected category and search query
   */
  private _getFilteredItems(): IFaqItem[] {
    let filtered: IFaqItem[] = [];

    // First filter by category
    const filterCategory = this.properties.filterCategory;

    if (this._selectedCategory === 'All') {
      if (!filterCategory) {
        filtered = this._faqItems.slice();
      } else {
        let i = 0;
        while (i < this._faqItems.length) {
          if (this._faqItems[i].category === filterCategory) {
            filtered.push(this._faqItems[i]);
          }
          i++;
        }
      }
    } else if (this._selectedCategory === this._topFaqCategoryKey) {
      let idx = 0;
      while (idx < this._faqItems.length) {
        const item = this._faqItems[idx];
        // Apply filter category if exists
        if (filterCategory && item.category !== filterCategory) {
          idx++;
          continue;
        }

        if (item.isTopFaq) {
          filtered.push(this._faqItems[idx]);
        }
        idx++;
      }
    } else {
      let i = 0;
      while (i < this._faqItems.length) {
        const item = this._faqItems[i];

        // Check main category filter first
        if (filterCategory && item.category !== filterCategory) {
          i++;
          continue;
        }

        // Check tab selection (Sub-Category if filtered, otherwise Category)
        const itemCategory = filterCategory ? item.subCategory : item.category;

        if (itemCategory === this._selectedCategory) {
          filtered.push(item);
        }
        i++;
      }
    }

    // Then filter by search query
    if (this._searchQuery && this._searchQuery.trim() !== '') {
      const query = this._searchQuery.toLowerCase().trim();
      const searchResults: IFaqItem[] = [];

      let j = 0;
      while (j < filtered.length) {
        const item = filtered[j];
        const titleMatch = item.title && item.title.toLowerCase().indexOf(query) !== -1;
        let answerMatch = false;

        if (this.properties.searchInAnswers && item.description) {
          // Strip HTML tags for search using DOM-based sanitization
          const plainDescription = FaqTextHelper.stripHtmlTags(item.description);
          answerMatch = plainDescription.toLowerCase().indexOf(query) !== -1;
        }

        if (titleMatch || answerMatch) {
          searchResults.push(item);
        }
        j++;
      }

      filtered = searchResults;
    }

    return filtered;
  }

  /**
   * Get total count of items (before search filter, after category filter)
   */
  private _getTotalItemsInCategory(): number {
    const filterCategory = this.properties.filterCategory;

    if (this._selectedCategory === 'All') {
      if (!filterCategory) {
        return this._faqItems.length;
      }
      // Count items matching filterCategory
      let count = 0;
      let i = 0;
      while (i < this._faqItems.length) {
        if (this._faqItems[i].category === filterCategory) {
          count++;
        }
        i++;
      }
      return count;
    }

    if (this._selectedCategory === this._topFaqCategoryKey) {
      let topCount = 0;
      let idx = 0;
      while (idx < this._faqItems.length) {
        const item = this._faqItems[idx];
        // Apply filter category if exists
        if (filterCategory && item.category !== filterCategory) {
          idx++;
          continue;
        }

        if (item.isTopFaq) {
          topCount++;
        }
        idx++;
      }
      return topCount;
    }

    let count = 0;
    let i = 0;
    while (i < this._faqItems.length) {
      const item = this._faqItems[i];

      // Check main category filter first
      if (filterCategory && item.category !== filterCategory) {
        i++;
        continue;
      }

      const itemCategory = filterCategory ? item.subCategory : item.category;

      if (itemCategory === this._selectedCategory) {
        count++;
      }
      i++;
    }
    return count;
  }

  private _getFontSize(size: string | undefined, defaultSize: string): string {
    return (size || defaultSize) + 'px';
  }

  private _shouldHighlightResults(): boolean {
    return !!(this.properties.highlightSearchMatches && this._searchQuery && this._searchQuery.trim() !== '');
  }

  private _prepareDeepLinkContext(): void {
    if (this._pendingDeepLinkId === null) {
      return;
    }

    const targetItem = this._faqItems.find((item: IFaqItem) => item.id === this._pendingDeepLinkId);
    if (!targetItem) {
      return;
    }

    if (this._searchQuery) {
      this._searchQuery = '';
    }

    if (this.properties.categoryColumn && targetItem.category) {
      // If using sub-category tabs, we need to match the sub-category
      if (this.properties.filterCategory) {
        if (targetItem.category === this.properties.filterCategory && targetItem.subCategory) {
          this._selectedCategory = targetItem.subCategory;
          return;
        }
        // If item not in filtered category, we can't show it in this focused view easily without complex logic
        // For now, fallback to 'All' or just let it be found via search if user clears filter
      } else {
        this._selectedCategory = targetItem.category;
        return;
      }
    }

    if (this._isTopFaqEnabled() && targetItem.isTopFaq) {
      this._selectedCategory = this._topFaqCategoryKey;
      return;
    }

    this._selectedCategory = 'All';
  }

  private _setPendingDeepLinkFromHash(hash: string | undefined): void {
    const targetId = this._getFaqIdFromHash(hash);
    this._pendingDeepLinkId = targetId;
    this._activeDeepLinkId = targetId;
  }

  private _getFaqIdFromHash(hash: string | undefined): number | null {
    if (!hash || hash.trim() === '') {
      return null;
    }

    const match = hash.trim().match(/^#faq-(\d+)$/i);
    if (!match) {
      return null;
    }

    const parsed = Number.parseInt(match[1], 10);
    return Number.isNaN(parsed) ? null : parsed;
  }

  private _handleHashChange = (): void => {
    if (typeof window === 'undefined') {
      return;
    }

    this._setPendingDeepLinkFromHash(window.location.hash);
    this.render();
  };

  /**
   * Render the web part
   */
  public render(): void {
    try {
      const semanticColors = this._themeVariant ? this._themeVariant.semanticColors : undefined;
      const palette = this._themeVariant ? this._themeVariant.palette : undefined;

      const headerBgColor = (palette && palette.themePrimary) ? palette.themePrimary : '#0078d4';
      const headerBgColorDark = (palette && palette.themeDark) ? palette.themeDark : '#005a9e';
      const bodyBackground = (semanticColors && semanticColors.bodyBackground) ? semanticColors.bodyBackground : '#ffffff';
      const bodyText = (semanticColors && semanticColors.bodyText) ? semanticColors.bodyText : '#323130';
      const bodySubtext = (semanticColors && semanticColors.bodySubtext) ? semanticColors.bodySubtext : '#605e5c';
      const linkColor = (semanticColors && semanticColors.link) ? semanticColors.link : ((palette && palette.themePrimary) ? palette.themePrimary : '#0078d4');
      const bodyDivider = (semanticColors && semanticColors.bodyDivider) ? semanticColors.bodyDivider : '#edebe9';
      const listItemBackgroundHovered = (semanticColors && semanticColors.listItemBackgroundHovered) ? semanticColors.listItemBackgroundHovered : '#f3f2f1';
      const cardBackground = (semanticColors && semanticColors.cardStandoutBackground) ? semanticColors.cardStandoutBackground : bodyBackground;
      const neutralLighter = (palette && palette.neutralLighter) ? palette.neutralLighter : '#f4f4f4';
      const neutralTertiary = (palette && palette.neutralTertiary) ? palette.neutralTertiary : '#a6a6a6';
      const neutralLight = (palette && palette.neutralLight) ? palette.neutralLight : '#eaeaea';
      const inputBackground = (semanticColors && semanticColors.inputBackground) ? semanticColors.inputBackground : '#ffffff';
      const inputBorder = (semanticColors && semanticColors.inputBorder) ? semanticColors.inputBorder : '#8a8886';

      const webpartTitleFontSize = this._getFontSize(this.properties.webpartTitleFontSize, '24');
      const webpartDescriptionFontSize = this._getFontSize(this.properties.webpartDescriptionFontSize, '14');
      const faqTitleFontSize = this._getFontSize(this.properties.faqTitleFontSize, '16');
      const faqDescriptionFontSize = this._getFontSize(this.properties.faqDescriptionFontSize, '14');
      const tabFontSize = this._getFontSize(this.properties.tabFontSize, '14');

      const searchPlaceholder = this.properties.searchPlaceholder || 'Search FAQs...';
      this._prepareDeepLinkContext();
      const shouldHighlight = this._shouldHighlightResults();
      const highlightQuery = shouldHighlight ? this._searchQuery.trim() : '';

      let faqItemsHtml = '';
      let categoryTabsHtml = '';
      let searchHtml = '';
      let searchInHeaderHtml = '';
      let resultsInfoHtml = '';

      // Build search input HTML
      const searchInputHtml = '<div class="' + styles.searchContainer + '">' +
        '<svg class="' + styles.searchIcon + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
        '<circle cx="11" cy="11" r="8"/>' +
        '<path d="M21 21l-4.35-4.35"/>' +
        '</svg>' +
        '<input type="text" class="' + styles.searchInput + '" ' +
        'placeholder="' + this._escapeHtml(searchPlaceholder) + '" ' +
        'value="' + this._escapeHtml(this._searchQuery) + '" ' +
        'style="background-color: ' + inputBackground + '; border-color: ' + inputBorder + '; color: ' + bodyText + ';">' +
        (this._searchQuery ? '<button class="' + styles.clearButton + '" title="Clear search" style="color: ' + bodySubtext + ';">' +
          '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
          '<line x1="18" y1="6" x2="6" y2="18"/>' +
          '<line x1="6" y1="6" x2="18" y2="18"/>' +
          '</svg>' +
          '</button>' : '') +
        '</div>';

      if (!this.properties.selectedList) {
        faqItemsHtml = '<div class="' + styles.emptyState + '" style="color: ' + bodySubtext + ';">' +
          '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
          '<circle cx="12" cy="12" r="10"/>' +
          '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>' +
          '<line x1="12" y1="17" x2="12.01" y2="17"/>' +
          '</svg>' +
          '<p>Please configure the web part by selecting a list from the property pane.</p>' +
          '</div>';
      } else {
        // Build search section (below header)
        if (this.properties.enableSearch && !this.properties.searchInHeader) {
          searchHtml = '<div class="' + styles.searchSection + '" style="background-color: ' + neutralLighter + '; border-bottom-color: ' + bodyDivider + ';">' +
            searchInputHtml +
            '</div>';
        }

        // Build search in header
        if (this.properties.enableSearch && this.properties.searchInHeader) {
          searchInHeaderHtml = '<div class="' + styles.searchInHeader + '">' + searchInputHtml + '</div>';
        }

        // Build category tabs (All + optional Top FAQ + categories)
        // If filtering by category, tabs are based on sub-category column.
        // We still check categoryColumn because it's the primary switch for tabs
        const hasCategoryTabs = !!(this.properties.categoryColumn && this._categories.length > 0);
        const showTopFaqTab = this._isTopFaqEnabled();
        if (hasCategoryTabs || showTopFaqTab) {
          const tabsArray: string[] = [];

          const allActiveClass = this._selectedCategory === 'All' ? ' ' + styles.activeTab : '';
          tabsArray.push(
            '<button class="' + styles.categoryTab + allActiveClass + '" ' +
            'data-category="All" ' +
            'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
            'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + '; ' +
            'font-size: ' + tabFontSize + ';">' +
            'All' +
            '</button>'
          );

          if (showTopFaqTab) {
            const isTopActive = this._selectedCategory === this._topFaqCategoryKey;
            const topActiveClass = isTopActive ? ' ' + styles.activeTab : '';
            const topFaqColor = this._getTopFaqColor(headerBgColor);
            const topBadgePalette = this._getTopFaqBadgePalette(topFaqColor);
            const badgeStyle = ' style="--topFaqBadgeBackground: ' + this._escapeHtml(topBadgePalette.badgeBackground) +
              '; --topFaqBadgeColor: ' + this._escapeHtml(topBadgePalette.badgeTextColor) +
              '; --topFaqBadgeBorder: ' + this._escapeHtml(topBadgePalette.badgeBorder) +
              '; --topFaqBadgeShadow: ' + this._escapeHtml(topBadgePalette.badgeShadow) + ';"';
            tabsArray.push(
              '<button class="' + styles.categoryTab + topActiveClass + '" ' +
              'data-category="' + this._topFaqCategoryKey + '" ' +
              'style="color: ' + (isTopActive ? headerBgColor : bodyText) + '; ' +
              'border-bottom-color: ' + (isTopActive ? headerBgColor : 'transparent') + '; ' +
              'font-size: ' + tabFontSize + ';">' +
              '<span class="' + styles.topFaqTab + '">' +
              '<span class="' + styles.topFaqBadge + '"' + badgeStyle + ' style="font-size: 1.2em;">⭐</span>' +
              '<span class="' + styles.topFaqBadgeLabel + '" style="font-size: 1em;">Top FAQ</span>' +
              '</span>' +
              '</button>'
            );
          }

          if (hasCategoryTabs) {
            let catIdx = 0;
            while (catIdx < this._categories.length) {
              const category = this._categories[catIdx];
              const isActive = this._selectedCategory === category;
              const activeClass = isActive ? ' ' + styles.activeTab : '';
              tabsArray.push(
                '<button class="' + styles.categoryTab + activeClass + '" ' +
                'data-category="' + this._escapeHtml(category) + '" ' +
                'style="color: ' + (isActive ? headerBgColor : bodyText) + '; ' +
                'border-bottom-color: ' + (isActive ? headerBgColor : 'transparent') + '; ' +
                'font-size: ' + tabFontSize + ';">' +
                this._escapeHtml(category) +
                '</button>'
              );
              catIdx++;
            }
          }

          categoryTabsHtml = '<div class="' + styles.categoryTabs + '" style="border-bottom-color: ' + neutralLight + ';">' +
            tabsArray.join('') +
            '</div>';
        }

        // Get filtered items
        const filteredItems = this._getFilteredItems();
        const totalItems = this._getTotalItemsInCategory();

        // Build results info
        if (this.properties.enableSearch && this.properties.showResultsCount && this._searchQuery && this._searchQuery.trim() !== '') {
          resultsInfoHtml = '<div class="' + styles.resultsInfo + '" style="color: ' + bodySubtext + ';">' +
            'Showing <strong style="color: ' + bodyText + ';">' + filteredItems.length + '</strong> of ' +
            '<strong style="color: ' + bodyText + ';">' + totalItems + '</strong> FAQs matching ' +
            '"<strong style="color: ' + bodyText + ';">' + this._escapeHtml(this._searchQuery) + '</strong>"' +
            '</div>';
        }

        if (filteredItems.length === 0) {
          const noResultsMessage = this._searchQuery && this._searchQuery.trim() !== ''
            ? 'No FAQs found matching "' + this._escapeHtml(this._searchQuery) + '". Try different keywords or clear the search.'
            : 'No FAQ items found' + (this._selectedCategory !== 'All' ? ' in this category' : '') + '.';

          faqItemsHtml = '<div class="' + styles.emptyState + '" style="color: ' + bodySubtext + ';">' +
            '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
            (this._searchQuery ?
              '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/><line x1="8" y1="11" x2="14" y2="11"/>' :
              '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>'
            ) +
            '</svg>' +
            '<p>' + noResultsMessage + '</p>' +
            '</div>';
        } else {
          const itemsHtmlArray: string[] = [];
          let index = 0;
          while (index < filteredItems.length) {
            const item = filteredItems[index];
            let attachmentsHtml = '';

            if (item.attachments && item.attachments.length > 0) {
              const attachmentLinksArray: string[] = [];
              let attIdx = 0;
              while (attIdx < item.attachments.length) {
                const att = item.attachments[attIdx];
                attachmentLinksArray.push(
                  '<a href="' + att.url + '" target="_blank" rel="noopener noreferrer" class="' + styles.attachmentLink + '" style="color: ' + linkColor + ';">' +
                  '<span class="' + styles.attachmentIcon + '">' +
                  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                  '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                  '<polyline points="14 2 14 8 20 8"/>' +
                  '</svg>' +
                  '</span>' +
                  this._escapeHtml(att.fileName) +
                  '</a>'
                );
                attIdx++;
              }

              attachmentsHtml = '<div class="' + styles.attachments + '" style="background-color: ' + neutralLighter + ';">' +
                '<div class="' + styles.attachmentsLabel + '" style="color: ' + neutralTertiary + ';">' +
                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                '<path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>' +
                '</svg>' +
                ' Attachments' +
                '</div>' +
                attachmentLinksArray.join('') +
                '</div>';
            }

            // Apply highlighting to title
            const escapedTitle = this._escapeHtml(item.title);
            const highlightedTitle = shouldHighlight
              ? FaqTextHelper.highlightText(escapedTitle, highlightQuery, styles.searchHighlight, false)
              : escapedTitle;

            // Apply highlighting to description
            const formattedDescription = this._formatDescription(item.description);
            const highlightedDescription = this.properties.searchInAnswers && shouldHighlight
              ? FaqTextHelper.highlightText(
                formattedDescription,
                highlightQuery,
                styles.searchHighlight,
                formattedDescription.indexOf('<') !== -1
              )
              : formattedDescription;

            const faqDomId = 'faq-' + item.id;
            const isDeepLinked = this._activeDeepLinkId === item.id;
            const deepLinkClass = isDeepLinked ? ' ' + styles.deepLinked : '';
            itemsHtmlArray.push(
              '<div class="' + styles.faqItem + deepLinkClass + '" data-index="' + index + '" data-faq-id="' + item.id + '" id="' + faqDomId + '">' +
              '<div class="' + styles.faqQuestion + '" style="border-bottom-color: ' + bodyDivider + ';" data-hover-bg="' + listItemBackgroundHovered + '">' +
              '<div class="' + styles.faqQuestionInner + '">' +
              '<span class="' + styles.faqQuestionText + '" style="color: ' + bodyText + '; font-size: ' + faqTitleFontSize + ';">' +
              highlightedTitle +
              '</span>' +
              '<button type="button" class="' + styles.copyLinkButton + '" data-item-id="' + item.id + '" aria-label="Copy link to ' + this._escapeHtml(item.title) + '" title="Copy link">' +
              '<span class="' + styles.copyLinkIcon + '">' +
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
              '<path d="M10 13a5 5 0 0 0 7.54.54l2.92-2.92a5 5 0 0 0-7.07-7.07l-1.29 1.29"/>' +
              '<path d="M14 11a5 5 0 0 0-7.54-.54l-2.92 2.92a5 5 0 0 0 7.07 7.07l1.29-1.29"/>' +
              '</svg>' +
              '</span>' +
              '<span class="' + styles.copyLinkText + '" data-default-text="Copy link">Copy link</span>' +
              '</button>' +
              '</div>' +
              '<span class="' + styles.faqChevron + '" style="color: ' + bodySubtext + ';">' +
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
              '<path d="M6 9l6 6 6-6"/>' +
              '</svg>' +
              '</span>' +
              '</div>' +
              '<div class="' + styles.faqAnswer + '">' +
              '<div class="' + styles.faqAnswerContent + '" style="color: ' + bodySubtext + '; font-size: ' + faqDescriptionFontSize + ';">' +
              highlightedDescription +
              '</div>' +
              attachmentsHtml +
              '</div>' +
              '</div>'
            );
            index++;
          }
          faqItemsHtml = itemsHtmlArray.join('');
        }
      }

      // Build header section
      let headerHtml = '';
      if (this.properties.webpartTitle || this.properties.webpartDescription || searchInHeaderHtml) {
        headerHtml = '<div class="' + styles.faqHeader + '" style="background: linear-gradient(135deg, ' + headerBgColor + ' 0%, ' + headerBgColorDark + ' 100%);">';
        if (this.properties.webpartTitle) {
          headerHtml += '<h2 class="' + styles.faqTitle + '" style="font-size: ' + webpartTitleFontSize + ';">' + this._escapeHtml(this.properties.webpartTitle) + '</h2>';
        }
        if (this.properties.webpartDescription) {
          headerHtml += '<p class="' + styles.faqDescription + '" style="font-size: ' + webpartDescriptionFontSize + ';">' + this._escapeHtml(this.properties.webpartDescription) + '</p>';
        }
        headerHtml += searchInHeaderHtml;
        headerHtml += '</div>';
      }

      this.domElement.innerHTML = '<div class="' + styles.customFaq + '" style="background-color: ' + cardBackground + ';">' +
        headerHtml +
        searchHtml +
        resultsInfoHtml +
        categoryTabsHtml +
        '<div class="' + styles.faqList + '">' +
        faqItemsHtml +
        '</div>' +
        '</div>';

      this._attachEventListeners();
      this._attachTabEventListeners();
      this._attachSearchEventListeners();
      this._attachCopyLinkEventListeners();
      this._expandPendingFaqItem();
    } catch (error) {
      console.error('Error during render:', error);
      this.domElement.innerHTML = '<div style="padding: 20px; color: red;">Error rendering web part. Please check console for details.</div>';
    }
  }

  /**
   * Attach search event listeners
   */
  private _attachSearchEventListeners(): void {
    const searchInput = this.domElement.querySelector('.' + styles.searchInput) as HTMLInputElement;
    const clearButton = this.domElement.querySelector('.' + styles.clearButton) as HTMLButtonElement;
    const self = this;

    if (searchInput) {
      // Debounce search input
      let debounceTimer: number | undefined;

      searchInput.addEventListener('input', function (): void {
        if (debounceTimer) {
          clearTimeout(debounceTimer);
        }
        debounceTimer = setTimeout(function (): void {
          self._searchQuery = searchInput.value;
          self.render();
        }, 300) as unknown as number;
      });

      // Handle Escape key
      searchInput.addEventListener('keydown', function (e: KeyboardEvent): void {
        if (e.key === 'Escape') {
          self._searchQuery = '';
          self.render();
        }
      });

      // Focus the input if it had a value (maintain focus after re-render)
      if (this._searchQuery) {
        searchInput.focus();
        searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
      }
    }

    if (clearButton) {
      clearButton.addEventListener('click', function (): void {
        self._searchQuery = '';
        self.render();
      });
    }
  }

  private _attachCopyLinkEventListeners(): void {
    const copyButtons = this.domElement.querySelectorAll('.' + styles.copyLinkButton);
    const self = this;

    let idx = 0;
    while (idx < copyButtons.length) {
      const button = copyButtons[idx] as HTMLButtonElement;
      (function (btn: HTMLButtonElement): void {
        btn.addEventListener('click', function (event: MouseEvent): void {
          event.stopPropagation();
          event.preventDefault();
          const idAttr = btn.dataset.itemId;
          if (!idAttr) {
            return;
          }
          const itemId = Number.parseInt(idAttr, 10);
          if (!Number.isNaN(itemId)) {
            self._copyLinkToFaq(itemId, btn);
          }
        });
      })(button);
      idx++;
    }
  }

  private _copyLinkToFaq(itemId: number, trigger: HTMLButtonElement): void {
    const link = this._buildFaqLink(itemId);
    const handleSuccess = (): void => {
      this._applyCopyFeedback(trigger, true);
      this._updateHashSilently(itemId);
    };
    const handleFailure = (): void => {
      this._applyCopyFeedback(trigger, false);
    };

    if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(link)
        .then(handleSuccess)
        .catch(() => {
          const fallbackSuccess = this._fallbackCopy(link);
          if (fallbackSuccess) {
            handleSuccess();
          } else {
            handleFailure();
          }
        });
    } else {
      const fallbackSuccess = this._fallbackCopy(link);
      if (fallbackSuccess) {
        handleSuccess();
      } else {
        handleFailure();
      }
    }
  }

  private _fallbackCopy(value: string): boolean {
    try {
      const textarea = document.createElement('textarea');
      textarea.value = value;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      textarea.style.pointerEvents = 'none';
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const success = document.execCommand('copy');
      document.body.removeChild(textarea);
      return success;
    } catch (error) {
      console.warn('Unable to copy link using fallback:', error);
      return false;
    }
  }

  private _applyCopyFeedback(button: HTMLButtonElement, success: boolean): void {
    const textElement = button.querySelector('.' + styles.copyLinkText) as HTMLElement;
    if (textElement) {
      const defaultText = textElement.getAttribute('data-default-text') || 'Copy link';
      textElement.textContent = success ? 'Copied!' : 'Failed';
      window.setTimeout(() => {
        textElement.textContent = defaultText;
      }, 2000);
    }

    if (success) {
      button.classList.add(styles.copySuccess);
      window.setTimeout(() => {
        button.classList.remove(styles.copySuccess);
      }, 2000);
    }
  }

  private _buildFaqLink(itemId: number): string {
    if (typeof window === 'undefined') {
      return '#faq-' + itemId;
    }
    return window.location.origin + window.location.pathname + window.location.search + '#faq-' + itemId;
  }

  private _updateHashSilently(itemId: number): void {
    if (typeof window === 'undefined') {
      return;
    }
    const hash = '#faq-' + itemId;
    if (window.location.hash === hash) {
      return;
    }

    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, document.title, window.location.pathname + window.location.search + hash);
    } else {
      window.location.hash = hash;
    }
  }

  private _expandPendingFaqItem(): void {
    if (this._pendingDeepLinkId === null) {
      return;
    }

    const selector = '[data-faq-id="' + this._pendingDeepLinkId + '"]';
    const target = this.domElement.querySelector(selector) as HTMLElement;
    if (!target) {
      return;
    }

    if (!target.classList.contains(styles.expanded)) {
      target.classList.add(styles.expanded);
    }
    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    this._pendingDeepLinkId = null;
  }

  private _attachTabEventListeners(): void {
    const tabs = this.domElement.querySelectorAll('.' + styles.categoryTab);
    const self = this;

    let tabIdx = 0;
    while (tabIdx < tabs.length) {
      const tab = tabs[tabIdx] as HTMLElement;
      (function (t: HTMLElement): void {
        t.addEventListener('click', function (): void {
          const category = t.dataset.category;
          if (category) {
            self._selectedCategory = category;
            self.render();
          }
        });
      })(tab);
      tabIdx++;
    }
  }

  private _attachEventListeners(): void {
    const faqItems = this.domElement.querySelectorAll('.' + styles.faqItem);
    const questions = this.domElement.querySelectorAll('.' + styles.faqQuestion);
    const self = this;

    let index = 0;
    while (index < questions.length) {
      const questionElement = questions[index] as HTMLElement;

      (function (qEl: HTMLElement, idx: number): void {
        qEl.addEventListener('mouseenter', function (): void {
          const hoverBg = qEl.dataset.hoverBg;
          if (hoverBg) {
            qEl.style.backgroundColor = hoverBg;
          }
        });

        qEl.addEventListener('mouseleave', function (): void {
          qEl.style.backgroundColor = '';
        });

        qEl.addEventListener('click', function (): void {
          const faqItem = faqItems[idx];

          if (!self.properties.allowMultipleExpanded) {
            let i = 0;
            while (i < faqItems.length) {
              if (i !== idx && faqItems[i].classList.contains(styles.expanded)) {
                faqItems[i].classList.remove(styles.expanded);
              }
              i++;
            }
          }

          if (faqItem.classList.contains(styles.expanded)) {
            faqItem.classList.remove(styles.expanded);
          } else {
            faqItem.classList.add(styles.expanded);
          }
        });
      })(questionElement, index);

      index++;
    }
  }

  private _escapeHtml(text: string): string {
    if (!text) {
      return '';
    }
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  private _formatDescription(description: string): string {
    if (!description) {
      return '';
    }

    // Check if content appears to contain HTML
    const hasOpenTag = description.indexOf('<') !== -1;
    const hasCloseTag = description.indexOf('>') !== -1;

    if (hasOpenTag && hasCloseTag) {
      // Content has HTML, return as-is (will be rendered as HTML)
      return description;
    }

    // Plain text: escape HTML entities and convert newlines
    return this._escapeHtml(description).split('\n').join('<br>');
  }

  private _loadLists(): Promise<void> {
    return this._spService.getLists()
      .then((lists: IListInfo[]) => {
        this._lists = lists;
      })
      .catch((error: Error) => {
        console.error('Error loading lists:', error);
        this._lists = [];
      });
  }

  private _loadColumns(): Promise<void> {
    if (!this.properties.selectedList) {
      this._columns = [];
      return Promise.resolve();
    }

    return this._spService.getListColumns(this.properties.selectedList)
      .then((columns: IColumnInfo[]) => {
        this._columns = columns;
      })
      .catch((error: Error) => {
        console.error('Error loading columns:', error);
        this._columns = [];
      });
  }

  private _loadFaqItems(): Promise<void> {
    if (!this.properties.selectedList || !this.properties.titleColumn || !this.properties.descriptionColumn) {
      this._faqItems = [];
      return Promise.resolve();
    }

    return this._spService.getListItems(
      this.properties.selectedList,
      this.properties.titleColumn,
      this.properties.descriptionColumn,
      this.properties.categoryColumn || undefined,
      this.properties.subCategoryColumn || undefined,
      this.properties.topFaqColumn || undefined
    )
      .then((items: IFaqItem[]) => {
        this._faqItems = items;
        this._extractCategories();
        const hasTopFaqItems = this._isTopFaqEnabled() && this._hasTopFaqItems();
        if (hasTopFaqItems && this._selectedCategory === 'All' && this._pendingDeepLinkId === null) {
          this._selectedCategory = this._topFaqCategoryKey;
        }
        this._ensureValidCategorySelection();
      })
      .catch((error: Error) => {
        console.error('Error loading FAQ items:', error);
        this._faqItems = [];
        this._categories = [];
      });
  }

  protected onDispose(): void {
    if (typeof window !== 'undefined') {
      window.removeEventListener('hashchange', this._handleHashChange);
    }
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | boolean, newValue: string | boolean): void {
    const self = this;

    if (propertyPath === 'selectedList' && newValue !== oldValue) {
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';
      this.properties.categoryColumn = '';
      this.properties.subCategoryColumn = '';
      this.properties.filterCategory = '';
      this.properties.topFaqColumn = '';
      this._faqItems = [];
      this._categories = [];
      this._selectedCategory = 'All';
      this._searchQuery = '';

      this._loadColumns().then(function (): void {
        self.context.propertyPane.refresh();
        self.render();
      });
    } else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn' || propertyPath === 'subCategoryColumn' || propertyPath === 'topFaqColumn') && newValue !== oldValue) {
      this._loadFaqItems().then(function (): void {
        self.render();
      });
    } else if (propertyPath === 'filterCategory') {
      this._extractCategories();
      this._selectedCategory = 'All'; // Reset selection when filter changes
      this.render();
    } else {
      this.render();
    }
  }

  private _getListOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a list --' }
    ];

    let i = 0;
    while (i < this._lists.length) {
      options.push({ key: this._lists[i].id, text: this._lists[i].title });
      i++;
    }

    return options;
  }

  private _getTitleColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
    }

    return options;
  }

  private _getDescriptionColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        const typeLabel = col.type === 'Note' ? 'Multi-line' : 'Single-line';
        options.push({ key: col.internalName, text: col.title + ' (' + typeLabel + ')' });
      }
      i++;
    }

    return options;
  }

  private _getCategoryColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No category (disable tabs) --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Choice') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
    }

    return options;
  }

  private _getTopFaqColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No Top FAQ column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Boolean' || col.type === 'Choice' || col.type === 'Text') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
    }

    return options;
  }

  private _getFilterCategoryOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No Filter (Show All) --' }
    ];

    // Using _extractCategories logic but on raw data would be better, 
    // but here we can just use the categories we've already extracted if no filter was applied,
    // OR we need to scan all items. 
    // simpler approach: scan all loaded items for unique values in 'categoryColumn'

    if (!this._faqItems || this._faqItems.length === 0) {
      return options;
    }

    const uniqueCategories: { [key: string]: boolean } = {};
    const catArray: string[] = [];

    // Note: We need the RAW category values, but _faqItems might already be loaded.
    // However, if a filter is applied, _faqItems only contains filtered items? 
    // No, _loadFaqItems loads ALL items from the list, filtering is done in memory (except list view threshold).
    // EXCEPT: _spService.getListItems gets everything. 
    // SO: This works.

    let i = 0;
    while (i < this._faqItems.length) {
      // Logic in _processItem ensures item.category is set from categoryColumn
      const cat = this._faqItems[i].category;
      if (cat && !uniqueCategories[cat]) {
        uniqueCategories[cat] = true;
        catArray.push(cat);
      }
      i++;
    }

    catArray.sort();

    let j = 0;
    while (j < catArray.length) {
      options.push({ key: catArray[j], text: catArray[j] });
      j++;
    }

    return options;
  }

  private _getFontSizeOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: '10', text: '10px - Extra Small' },
      { key: '12', text: '12px - Small' },
      { key: '14', text: '14px - Normal' },
      { key: '16', text: '16px - Medium' },
      { key: '18', text: '18px - Large' },
      { key: '20', text: '20px - Extra Large' },
      { key: '24', text: '24px - Heading' },
      { key: '28', text: '28px - Large Heading' },
      { key: '32', text: '32px - Extra Large Heading' }
    ];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the FAQ web part settings'
          },
          groups: [
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: 'Webpart Title',
                  placeholder: 'Enter a title for the FAQ section'
                }),
                PropertyPaneDropdown('webpartTitleFontSize', {
                  label: 'Webpart Title Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.webpartTitleFontSize || '24'
                }),
                PropertyPaneTextField('webpartDescription', {
                  label: 'Webpart Description',
                  placeholder: 'Enter a description',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneDropdown('webpartDescriptionFontSize', {
                  label: 'Webpart Description Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.webpartDescriptionFontSize || '14'
                }),
                PropertyPaneDropdown('tabFontSize', {
                  label: 'Tabs Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.tabFontSize || '14'
                }),
                PropertyPaneTextField('topFaqColor', {
                  label: 'Top FAQ Accent Color',
                  placeholder: '#0078d4 or red',
                  description: 'Leave blank to inherit the theme color from the page.'
                })
              ]
            },
            {
              groupName: 'FAQ Item Settings',
              groupFields: [
                PropertyPaneDropdown('faqTitleFontSize', {
                  label: 'FAQ Question Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.faqTitleFontSize || '16'
                }),
                PropertyPaneDropdown('faqDescriptionFontSize', {
                  label: 'FAQ Answer Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.faqDescriptionFontSize || '14'
                })
              ]
            },
            {
              groupName: 'Search Settings',
              groupFields: [
                PropertyPaneToggle('enableSearch', {
                  label: 'Enable Search',
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('searchInHeader', {
                  label: 'Search in Header',
                  onText: 'Yes (in header)',
                  offText: 'No (below header)',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('highlightSearchMatches', {
                  label: 'Highlight Search Matches',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('searchInAnswers', {
                  label: 'Search in Answers',
                  onText: 'Yes (titles & answers)',
                  offText: 'No (titles only)',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('showResultsCount', {
                  label: 'Show Results Count',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneTextField('searchPlaceholder', {
                  label: 'Search Placeholder Text',
                  placeholder: 'Search FAQs...',
                  disabled: !this.properties.enableSearch
                })
              ]
            },
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneDropdown('selectedList', {
                  label: 'Select List',
                  options: this._getListOptions()
                }),
                PropertyPaneDropdown('titleColumn', {
                  label: 'Title Column',
                  options: this._getTitleColumnOptions(),
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('descriptionColumn', {
                  label: 'Description Column',
                  options: this._getDescriptionColumnOptions(),
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('categoryColumn', {
                  label: 'Category Column (Main Filter)',
                  options: this._getCategoryColumnOptions(),
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('subCategoryColumn', {
                  label: 'Sub-Category Column (for tabs)',
                  options: this._getCategoryColumnOptions(), // Reuse same options logic (Text/Choice)
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('filterCategory', {
                  label: 'Filter to Category',
                  options: this._getFilterCategoryOptions(),
                  disabled: !this.properties.selectedList || !this.properties.categoryColumn
                }),
                PropertyPaneDropdown('topFaqColumn', {
                  label: 'Top FAQ Column (Yes/No)',
                  options: this._getTopFaqColumnOptions(),
                  disabled: !this.properties.selectedList
                })
              ]
            },
            {
              groupName: 'Accordion Behavior',
              groupFields: [
                PropertyPaneToggle('allowMultipleExpanded', {
                  label: 'Allow Multiple Items Expanded',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
