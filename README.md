# Building a Modern Intranet: A Comprehensive SharePoint Framework (SPFx) Solution

Discover the **Modern Intranet Kit**, a professional-grade suite of SharePoint Framework (SPFx) web parts designed to transform your employee portal into a high-engagement, modern workplace. This blog post dives into the solution overview, the underlying toolchain, and a detailed breakdown of all the web parts included in the suite.

---

## 🚀 Solution Overview

The Modern Intranet Kit is a modular branding and functionality suite that empowers organizations to design unique, engaging layouts in SharePoint Online. 

By deploying a single `.sppkg` file, the entire kit is installed at once, instantly providing your team with multiple premium tools available in the web part picker.

### Core Value Pillars:
- **🛠️ Flexible Modern Layouts**: Build a personalized experience by choosing the right web parts for your communication strategy.
- **📦 Single Solution Deployment**: One single installation for the entire suite.
- **🎨 Native Theming**: Every component automatically inherits the SharePoint site theme for seamless branding.
- **⚡ Cross-Site Connectivity**: Easily pull and aggregate data from any list across your tenant.
- **📌 Intelligent Prioritization**: Built-in logic for "Pinned" and "Active" statuses allows editors to promote high-value content.

---

## 🛠️ Toolchain and Technical Foundation

The solution is built on modern web development standards to ensure optimum performance, security, and maintainability.

- **Framework**: SharePoint Framework (SPFx) 1.22.0
- **Language**: TypeScript (~5.8.0) and React (17.0.1)
- **UI Toolkit**: Fluent UI React (`@fluentui/react` ^8.106.4)
- **Build Tooling**: Node.js (>= 22.14.0 < 23.0.0), Rush Stack Heft (`@rushstack/heft` 1.1.2), and Webpack
- **Engine**: SharePoint REST API (`spHttpClient`) & Microsoft Graph API
- **Styling**: Vanilla CSS with localized SCSS Modules and full CSS Custom Property integration for theme inheritance
- **Architecture**: Shared `SiteListService` and `ThemeService` for unified data and visual logic
- **Security**: 100% CSP compliance with zero external dependencies, native SharePoint, and Microsoft Graph authentication.

---

## 📦 Web Part Catalog and Features

Here is a detailed look at the 10 core web parts included in the suite, their key features, and configuration capabilities:

### 1. Banner Slider
*High-impact visual communication for the top of your homepage.*
- **Description**: A full-width, auto-rotating image carousel for featured news and strategic announcements.
- **Key Features**: Auto-Rotation (3-10s) with hover-to-pause logic, glassmorphism overlays for text readability, dynamic calls-to-action, and "Active" status filtering.
- **Configuration Properties**: Data Source (`siteUrl`, `listId`), Mapping (`title`, `description`, `image`, `active`, `buttonText`, `pageLink`), Settings (`autoRotateInterval`, `showCta`, `showTitle`), and Styling (`titleBarStyle`).

### 2. Events
*A clean, organized way to display what's happening in your organization.*
- **Description**: A multi-column grid layout for upcoming engagement opportunities and corporate events.
- **Key Features**: Flexible grid density (2, 3, or 4 columns), chronological intelligence to hide past events, priority "pinning", and an optional "View All" link.
- **Configuration Properties**: Data Source (`siteUrl`, `listId`), Mapping (`title`, `date`, `image`, `link`, `location`, `active`, `pinned`), Layout (`maxItems`, `itemsPerRow`), Action (`showViewAll`, `viewAllUrl`).

### 3. Highlights
*Modular content cards for news, initiatives, or spotlights.*
- **Description**: A card-based grid featuring rich imagery, clear titles, and descriptive text snippets.
- **Key Features**: Pre-optimized thumbnails utilizing SharePoint Image columns, priority/pinned item sorting, and synchronized header styling.
- **Configuration Properties**: Mapping (`title`, `description`, `bannerImage`, `link`, `pinned`), Display (`maxItems`, `columns` 2/3), Branding (`showTitle`, `showBackgroundBar`, `titleBarStyle`).

### 4. Quick Links
*Streamlined navigation tiles for your most-used resources.*
- **Description**: A minimalist grid of shortcut tiles with icons and labels for efficient information access.
- **Key Features**: High-density layout (up to 6 tiles per row), Fluent UI icon integration, and smart internal/external navigation handling.
- **Configuration Properties**: Mapping (`title`, `link`, `icon`, `pinned`), Layout (`columnsPerRow`), UX (`openInNewTab`).

### 5. Document Viewer
*A centralized hub for navigating organizational files and resources.*
- **Description**: A multi-level document explorer featuring category-based navigation, sub-category filtering, and integrated search.
- **Key Features**: Dual-level Navigation, Side Navigation vs Top Tabs display options, seamless SharePoint Web Viewer (`?web=1`) integration, and instant real-time search.
- **Configuration Properties**: Mapping (`categoryField`, `subCategoryField`, `descriptionField`, `pinnedField`), Logic (`enableSubCategory`, `categoryDisplayType`), Styling (`pageSize`, font sizes).

### 6. Calendar
*A unified hub for managing your organizational schedule.*
- **Description**: A feature-rich calendar with support for Day, Week, Month, and Year views.
- **Key Features**: Native SharePoint Events list integration, yearly grid or timeline modes, live current-time indicator, and clean contextual overlays for event details.
- **Configuration Properties**: Mapping (`title`, `startDate`, `endDate`, `location`), Settings (`defaultView`, `yearViewType`), Branding (`showTitle`, `showBackgroundBar`, `titleBarStyle`).

### 7. Employee Directory
*A dynamic, searchable directory for finding and connecting with colleagues.*
- **Description**: A comprehensive employee catalog supporting both Azure AD (Graph) and custom SharePoint Lists.
- **Key Features**: Dual data sources (Graph vs SP List), profile enrichment fields (About Me, Skills, etc.), List/Grid view toggling, and smart filtering (Department, Location).
- **Configuration Properties**: Source selection, Full attribute mapping, UI settings (`viewMode`, `pageSize`, `showFilters`, `showPagination`).

### 8. Employee Spotlight
*Recognize and highlight exceptional team members.*
- **Description**: A dynamic showcase for spotlighting specific employees via rich cards or an auto-rotating carousel.
- **Key Features**: Microsoft Graph People Picker integration for manual selection, Standard/Compact layout modes, and custom write-ups per person.
- **Configuration Properties**: Source (`SP List` vs `Manual`), Manual (`selectedUsers`, `commonDescription`), Automation (`autoRotateInterval`, `maxItems`), Styling (`layoutMode`, `titleBarStyle`).

### 9. New Joiners
*A warm welcome for your organization's newest members.*
- **Description**: A visually engaging slider or grid that highlights recent hires, featuring photos and welcome messages.
- **Key Features**: Dedicated horizontal "Strip" mode for sidebars, manual curation via Graph picker, and customized welcome messages.
- **Configuration Properties**: Source selection, Mapping (`name`, `photo`, `jobTitle`, `department`, `newJoinerText`), Layout (`List`/`Grid`/`Strip`, layout mode), Settings (`maxItems`).

### 10. FAQ
*Accordion-style information hub with search and filtering.*
- **Description**: A performance-optimized FAQ web part featuring category filtering and real-time search.
- **Key Features**: Instant client-side search, dynamic category pills for filtering, Single/Multi expand modes, and ordered content sorting.
- **Configuration Properties**: Mapping (`question`, `answer`, `category`, `order`), Settings (`showSearch`, `showCategoryFilter`, `allowMultipleOpen`, `expandFirstItem`), Branding (`showTitle`/`titleBarStyle`).

---

## 💎 Conclusion

The **Modern Intranet Kit** standardizes the intranet experience with robust, professional configurations, providing an incredible user experience while putting immense control into the hands of site authors. With common branding elements, standardized loading states, and robust SharePoint/Graph integration, it forms a cohesive and powerful digital workplace foundation.
