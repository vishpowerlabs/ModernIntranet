# Modern Intranet Kit: Professional SPFx Web Part Suite

The **Modern Intranet Kit** is a comprehensive, professional-grade solution comprising various high-performance web parts designed to build clean, modern, and high-engagement employee portals. 

---

## 🚀 The Solution: A Modular Branding Kit

The Modern Intranet Kit provides organizations with the flexibility to design their own unique layouts. By deploying a single solution, all web parts are simultaneously added to the SharePoint App Catalog, making them immediately available in the web part picker for use on any modern SPO page.

### Core Value Pillars:
- **🛠️ Flexible Modern Layouts**: Build a personalized intranet experience by choosing the specific web parts that fit your communication strategy.
- **📦 Single Solution Deployment**: When the `.sppkg` is deployed, the entire kit is installed at once, instantly providing your team with multiple premium tools.
- **🎨 Native Theming**: Every component automatically inherits the SharePoint site theme. No hardcoded colors—it looks perfect on every site, every time.
- **⚡ Cross-Site Connectivity**: Pull data from any list across your tenant by simply entering the Site URL.
- **📌 Intelligent Prioritization**: Built-in "Pinned" and "Active" status logic allows editors to promote high-value content to the top of any list.

---

## 📦 Web Part Catalog

### 1. Modern - Banner Slider
*High-impact visual communication for the top of your homepage.*
- **Description**: A full-width, auto-rotating image carousel for featured news and strategic announcements.
- **Key Features**:
    - Auto-rotate with configurable timing.
    - Dynamic CTA (Call to Action) buttons.
    - Glassmorphism overlays for readability on any background image.
- **Configuration Properties**:
    - `siteUrl` & `listId`: Data source location.
    - `titleColumn` & `descriptionColumn`: Content mapping.
    - `imageColumn`: Choice of Banner Image.
    - `activeColumn`: Filter for current vs. expired banners.
    - `autoRotateInterval`: Seconds between slides.
    - `showCta`: Toggle for the action button.

### 2. Modern - Events
*A clean, organized way to display what's happening in your organization.*
- **Description**: A multi-column grid layout for upcoming engagement opportunities.
- **Key Features**:
    - **Dynamic Grid**: Switch between 2, 3, or 4 columns per row.
    - **Smart Filtering**: Automatically hides past events; supports "Active" status filtering.
    - **Pinned Logic**: High-priority events float to the top regardless of chronological order.
    - **Branding**: Includes a custom title bar with theme-aware background accents.
- **Configuration Properties**:
    - `maxItems`: Limits total display (supports up to 12).
    - `itemsPerRow`: Layout density (2, 3, or 4).
    - `dateColumn`: Mapping for the event date/time.
    - `activeColumn`: Mapping for "Active" status.
    - `pinnedColumn`: Mapping for prioritization.
    - `showViewAll`: Optional link to a full calendar page.

### 3. Modern - Highlights
*Modular content cards for news, initiatives, or spotlights.*
- **Description**: A card-based grid featuring rich imagery, clear titles, and descriptive text snippets.
- **Key Features**:
    - **Thumbnail Support**: Optimized for the SharePoint Image column type.
    - **Flexible Layout**: Configurable 2 or 3 column density.
    - **Priority Order**: Supports "Pinned" items for administrative control over sorting.
- **Configuration Properties**:
    - `maxItems`: Control the "footprint" on the page.
    - `columns`: Columns per row (2 or 3).
    - `bannerImageColumn`: The card's hero image.
    - `pinnedColumn`: Ensures specific highlights stay at the top.
    - `showTitle` & `showBackgroundBar`: Branding controls.

### 4. Modern - QuickLinks
*Streamlined navigation tiles for your most-used resources.*
- **Description**: A minimalist grid of shortcut tiles with icons and labels for efficient information access.
- **Key Features**:
    - **High-Density**: Supports up to 6 columns per row for compact navigation.
    - **Icon Integration**: Supports standard Fluent UI/Office UI Fabric icons.
    - **New Tab Control**: Choice to open internal links in-place or external links in a new window.
- **Configuration Properties**:
    - `columnsPerRow`: Choose between 2, 3, 4, or 6 columns.
    - `iconColumn`: Map a Choice or Text field to Fluent UI icon names.
    - `pinnedColumn`: Prioritize critical tools/links at the front of the list.
    - `openInNewTab`: Global behavior for navigation.

### 5. Modern - Document Viewer
*A centralized hub for navigating organizational files and resources.*
- **Description**: A multi-level document explorer featuring category-based navigation, sub-category filtering, and integrated search.
- **Key Features**:
    - **Adaptive Navigation**: Supports dual-level filtering. Navigate by main Category and further refine by Sub-Category.
    - **Flexible Layouts**: Choose whether to display Categories in a professional **Side Navigation** panel or as **Top Tabs**.
    - **Sub-Category Control**: Toggle sub-categories on or off depending on the organizational depth of your library.
    - **Web Viewer Integration**: Documents open instantly in a new tab using the SharePoint Web Viewer (`?web=1`), ensuring users can view Office documents and PDFs without downloading.
    - **Integrated Search**: Real-time keyword filtering across document titles and descriptions.
    - **Prioritization Logic**: Highlights essential policies or forms by pinning them to the top of the list with a visual indicator.
    - **Premium UI**: Housed in a theme-aware card container with 18px titles and a dynamic accent bar matching other kit components.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId` (Connect to any library in the tenant).
    - **Column Mapping**: `categoryField`, `subCategoryField`, `descriptionField`, and `pinnedField`.
    - **Navigation Logic**: `enableSubCategory` (Toggle) and `categoryDisplayType` (Side Nav vs. Top Tabs).
    - **Display Controls**: `pageSize` (Items per page), `webPartTitle`, and `webPartDescription`.

---

## 🛠 Technical Foundation

- **Framework**: SPFx 1.22.0
- **Language**: TypeScript / React
- **API**: SharePoint REST API (spHttpClient)
- **Styling**: Vanilla CSS with SPFx CSS Modules
- **Theming**: CSS Custom Properties (`--bodyText`, `--themePrimary`, etc.)
- **Service Layer**: Dedicated `SiteListService` for robust data fetching across hierarchies.

---

## 🔒 Security & CSP Compliance

The Modern Intranet Kit is fully compatible with **SharePoint Content Security Policy (CSP)**. 

- **Trusted Assets**: All solution assets (JavaScript, CSS, and manifests) are hosted directly within your SharePoint environment (tenant-hosted or site-level). 
- **Zero External Dependencies**: By keeping all assets within the SharePoint trusted zone, the solution inherently adheres to strict CSP policies without requiring additional domain whitelisting.
- **Data Privacy**: All data fetching is performed via the native `spHttpClient`, ensuring that authentication and data transit remain securely within the user's SharePoint context.

### 6. Modern - Calendar
*A unified hub for managing your organizational schedule.*
- **Description**: A feature-rich calendar with support for Day, Week, Month, and Year views. Optimized for standard SharePoint Calendar lists.
- **Key Features**:
    - **Multi-View Navigation**: Seamlessly switch between micro-level daily timelines and high-level yearly overviews.
    - **Native List Integration**: Built-in support for standard SharePoint Events (Template 106).
    - **Live Timeline**: Real-time "current time" indicator in Day and Week views.
    - **Contextual Headers**: Smart date labels (e.g., "March 2026", "Week of Mar 3") based on active view.
    - **Smart Popups**: Clean, theme-aware popups for viewing event details without leaving the page.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Schema Mapping**: `titleColumn`, `dateColumn` (Start), `endDateColumn`, and `locationColumn`.
    - **Display & Branding**: `defaultView`, `showTitle`, `title` (Web Part Title), and `showBackgroundBar`.

---

> [!NOTE]
> This document is intended as a foundation for product marketing pages, technical documentation, and user guides.
