# Modern Intranet Kit: Professional SPFx Web Part Suite

The **Modern Intranet Kit** is a comprehensive, professional-grade solution comprising various high-performance web parts designed to build clean, modern, and high-engagement employee portals. 

---

## 🚀 The Solution: A Modular branding Kit

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
    - **Auto-Rotation**: Configurable timing (3-10s) with hover-to-pause logic.
    - **Glassmorphism Overlays**: Semi-transparent content cards ensure readability regardless of the background image.
    - **Dynamic Call-to-Action**: Integrated buttons for direct navigation to news articles or landing pages.
    - **Status Filtering**: Built-in logic to show only "Active" banners.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `titleColumn`, `descriptionColumn`, `imageColumn`, `activeColumn`, `buttonTextColumn`, `pageLinkColumn`.
    - **Settings**: `autoRotateInterval` (Seconds), `showCta` (Toggle), `showTitle` (Branding).
    - **Styling**: `titleBarStyle` (Solid/Underline).

### 2. Modern - Events
*A clean, organized way to display what's happening in your organization.*
- **Description**: A multi-column grid layout for upcoming engagement opportunities and corporate events.
- **Key Features**:
    - **Flexible Grid density**: Choose between 2, 3, or 4 columns per row to fit your layout.
    - **Chronological Intelligence**: Events are sorted by date, and past events can be automatically hidden.
    - **Pinned Priority**: Specific events can be "pinned" to stay at the beginning of the list.
    - **View All Integration**: Optional button to redirect users to a full calendar or event library.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `titleColumn`, `dateColumn`, `imageColumn`, `linkColumn`, `locationColumn`, `activeColumn`, `pinnedColumn`.
    - **Layout**: `maxItems` (Total count), `itemsPerRow` (2/3/4).
    - **Action**: `showViewAll` (Toggle), `viewAllUrl` (Link).

### 3. Modern - Highlights
*Modular content cards for news, initiatives, or spotlights.*
- **Description**: A card-based grid featuring rich imagery, clear titles, and descriptive text snippets.
- **Key Features**:
    - **Thumbnail Optimization**: Specifically designed to pull high-quality thumbnails from SharePoint Image columns.
    - **Priority Order**: Supports "Pinned" items for administrative control over sorting.
    - **Branded Headers**: Synchronized styling with the rest of the kit via the header accent bar.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `titleColumn`, `descriptionColumn`, `bannerImageColumn`, `linkColumn`, `pinnedColumn`.
    - **Display**: `maxItems` (3-12), `columns` (2 or 3 per row).
    - **Branding**: `showTitle`, `showBackgroundBar`, `titleBarStyle`.

### 4. Modern - Quick Links
*Streamlined navigation tiles for your most-used resources.*
- **Description**: A minimalist grid of shortcut tiles with icons and labels for efficient information access.
- **Key Features**:
    - **High-Density Layout**: Supports up to 6 tiles per row, making it perfect for "Toolbox" sections.
    - **Icon Integration**: Full support for Fluent UI (Office UI Fabric) icon sets.
    - **Smart Navigation**: Choice to open internal links in-place or launch external resources in new tabs.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `titleColumn`, `linkColumn`, `iconColumn`, `pinnedColumn`.
    - **Layout**: `columnsPerRow` (2, 3, 4, or 6).
    - **UX**: `openInNewTab` (Toggle).

### 5. Modern - Document Viewer
*A centralized hub for navigating organizational files and resources.*
- **Description**: A multi-level document explorer featuring category-based navigation, sub-category filtering, and integrated search.
- **Key Features**:
    - **Dual-Level Navigation**: Navigate by main Category and further refine by Sub-Category.
    - **UI Layout Selection**: Display Categories in a professional **Side Navigation** panel or as **Top Tabs**.
    - **Web Viewer Integration**: Documents open instantly using SharePoint Web Viewer (`?web=1`) for seamless Office/PDF viewing.
    - **Instant Search**: Real-time keyword filtering across document titles and descriptions.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `categoryField`, `subCategoryField`, `descriptionField`, `pinnedField`.
    - **Logic**: `enableSubCategory` (Toggle), `categoryDisplayType` (Side Nav vs Top Tabs).
    - **Styling**: `pageSize`, `webPartTitleFontSize`, `webPartDescriptionFontSize`, `headerOpacity`.

### 6. Modern - Calendar
*A unified hub for managing your organizational schedule.*
- **Description**: A feature-rich calendar with support for Day, Week, Month, and Year views.
- **Key Features**:
    - **Native SharePoint Integration**: Built-in support for standard SharePoint Events (Template 106) lists.
    - **Yearly View Modes**: Choose between a high-level **Grid** overview or a chronological **Timeline**.
    - **Micro-Timelines**: Live "current time" indicator in Day and Week views.
    - **Contextual Overlays**: Clean, theme-aware popups for viewing event details without site navigation.
- **Configuration Properties**:
    - **Data Source**: `siteUrl` & `listId`.
    - **Mapping**: `titleColumn`, `dateColumn` (Start), `endDateColumn`, `locationColumn`.
    - **Settings**: `defaultView` (Day/Week/Month/Year), `yearViewType` (Grid/Timeline).
    - **Branding**: `showTitle`, `showBackgroundBar`, `titleBarStyle`.

### 7. Modern - Employee Directory
*A dynamic, searchable directory for finding and connecting with colleagues.*
- **Description**: A comprehensive employee catalog supporting both Azure AD (Graph) and custom SharePoint Lists.
- **Key Features**:
    - **Dual Data Sources**: Seamlessly switch between the entire organization's Azure AD or a curated SharePoint List.
    - **Rich Profiles**: Extend employee cards with custom data like "About Me", "Projects", "Skills", and "Interests".
    - **View Switching**: Instantly toggle between a dense 'List View' and a visually rich 'Grid View' (profile cards).
    - **Smart Filtering**: Global search plus dynamic dropdown filters for Department and Location.
- **Configuration Properties**:
    - **Source**: `source` (Graph vs. SP List).
    - **Mapping**: Full mapping for Name, Photo, Job Title, Department, Email, Location, Manager, and Profile Enrichment fields.
    - **UI**: `viewMode` (List/Grid), `pageSize`, `showFilters`, `showPagination`.

### 8. Modern - Employee Spotlight
*Recognize and highlight exceptional team members.*
- **Description**: A dynamic showcase for spotlighting specific employees via rich cards or an auto-rotating carousel.
- **Key Features**:
    - **People Picker Integration**: Manually select employees via the Microsoft Graph People Picker for instant recognition.
    - **Layout Modes**: **Standard (Wide)** for main content areas or **Compact** for sidebars and secondary columns.
    - **Custom Write-ups**: Supports unique spotlight descriptions or a common global recognition message.
- **Configuration Properties**:
    - **Source**: `source` (SP List vs Manual Graph Selection).
    - **Manual**: `selectedUsers` (People Picker), `commonDescription`.
    - **Automation**: `autoRotateInterval` (3-15s), `maxItems`.
    - **Styling**: `layoutMode` (Standard/Compact), `webPartTitleFontSize`, `titleBarStyle`.

### 9. Modern - New Joiners
*A warm welcome for your organization's newest members.*
- **Description**: A visually engaging slider or grid that highlights recent hires, featuring photos and welcome messages.
- **Key Features**:
    - **Horizontal "Strip" Mode**: A premium, auto-rotating horizontal slider specifically designed for sidebar placements.
    - **Manual Curation**: Support for manually adding newcomers via Graph picker when lists aren't desired.
    - **Welcome Intro**: Supports individual tailored greetings or a global fallback message.
- **Configuration Properties**:
    - **Source**: `source` (SP List vs Manual).
    - **Mapping**: `nameColumn`, `photoColumn`, `jobTitleColumn`, `departmentColumn`, `newJoinerTextColumn`.
    - **UI**: `layout` (List/Grid/Strip), `layoutMode` (Standard/Compact), `maxItems`.

### 10. Modern - FAQ
*Accordion-style information hub with search and filtering.*
- **Description**: A performance-optimized FAQ web part featuring category filtering and real-time search.
- **Key Features**:
    - **Client-Side Search**: Instant keyword filtering with visual highlighting.
    - **Category Pills**: Dynamic filtering by topic to quickly find relevant answers.
    - **Interaction Logic**: Choose between "Single Expand" or "Multi Expand" accordion behaviors.
    - **Ordered Content**: Supports numerical order mapping to ensure critical questions appear first.
- **Configuration Properties**:
    - **Mapping**: `questionColumn`, `answerColumn`, `categoryColumn`, `orderColumn`.
    - **Settings**: `showSearch`, `showCategoryFilter`, `allowMultipleOpen`, `expandFirstItem`.
    - **Branding**: `showTitle`, `showBackgroundBar`, `titleBarStyle`.

---

## 💎 UI Excellence & Standardization

The Modern Intranet Kit is built for professional consistency. Every web part in the suite shares a harmonized user experience:

- **Standardized Loading States**: Uses centered, branded Fluent UI Spinners during all data fetching operations.
- **Professional Empty States**: Clear, actionable guidance with custom iconography when configuration is required.
- **Theming & Branding**: Every web part includes a customizable title container with 'Solid' or 'Underline' accent bars that automatically sync with the site's primary theme color.
- **Responsive Layout Modes**: Components support 'Standard' and 'Compact' modes, ensuring they look perfect in any page section.

---

## 🛠 Technical Foundation

- **Platform**: SPFx 1.22.0
- **Language**: TypeScript / React
- **Engine**: SharePoint REST API (spHttpClient) & Microsoft Graph
- **Styling**: Vanilla CSS with localized CSS Modules
- **Theming**: Full CSS Custom Property integration for theme inheritance
- **Architecture**: Shared `SiteListService` and `ThemeService` for unified data and visual logic.

---

## 🔒 Security & CSP Compliance

- **Trusted Execution**: All assets strictly hosted within the SharePoint tenant, ensuring 100% CSP compliance.
- **Zero External Dependencies**: No external CDNs or third-party scripts required.
- **Secure Data Access**: Uses native SharePoint and Graph authentication contexts for all data operations.

---

> [!NOTE]
> This document serves as the master source for technical specifications, product documentation, and deployment guides.
