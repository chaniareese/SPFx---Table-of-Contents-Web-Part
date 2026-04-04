# GEMINI.md - SPFx Table of Contents Web Part

## Project Overview
This project is a **SharePoint Framework (SPFx)** web part that provides a dynamic **Table of Contents (TOC)** and a rich-text content area. It is designed to help users create organized, long-form content (like runbooks or documentation) with an automatically generated navigation sidebar.

- **Primary Purpose:** Provide an integrated editor and navigation experience within SharePoint pages.
- **Main Technologies:**
  - **SPFx 1.22.2**: The core framework for the web part.
  - **React 17**: UI library for components.
  - **TipTap 3.20.1**: A headless rich-text editor used for content creation.
  - **Heft**: The build toolchain (part of the Rush stack).
  - **Fluent UI React**: For UI components and styling consistency.

## Architecture
The project follows a standard SPFx React-based architecture:

- **Web Part Wrapper (`TableOfContentsWebPart.ts`):** 
  - Manages property pane configuration (TOC title, visibility of H2/H3/H4 headings, mobile display).
  - Initializes the React component with web part properties.
- **Main Component (`TableOfContents.tsx`):**
  - Manages the state of the content and the extracted headings.
  - Implements scroll-sync logic to highlight the active heading in the TOC as the user scrolls.
  - Handles the layout, alternating between the sidebar (`TOCNavigator`) and the editor/display area.
- **Editor Component (`TipTapEditor.tsx`):**
  - A highly customized TipTap implementation.
  - Includes custom extensions for Font Size, Line Height, Indentation, Resizable Images, and advanced Table support.
  - Provides a floating toolbar for text and table formatting.
- **Navigator Component (`TOCNavigator.tsx`):**
  - Renders the hierarchical tree of headings (H2, H3, H4).
  - Handles smooth scrolling to sections when a link is clicked.

## Building and Running
The project uses `heft` for all build and lifecycle tasks.

- **Prerequisites:**
  - Node.js `^22.14.0` (as specified in `package.json`).
  - `@rushstack/heft` installed globally (`npm install -g @rushstack/heft`).

- **Key Commands:**
  - `heft start`: Starts the local development server and workbench.
  - `heft clean`: Cleans the build artifacts.
  - `heft test --clean --production`: Runs tests (if any) and cleans the production build.
  - `heft package-solution --production`: Packages the solution into a `.sppkg` file for deployment.

## Development Conventions
- **Styling:** Uses **SCSS Modules** (`*.module.scss`) for component-scoped styling.
- **TypeScript:** Strict typing is enforced. Custom TipTap commands are declared via module augmentation in `TipTapEditor.tsx`.
- **Property Management:** Web part properties are persisted in the `properties` object and synced with the React state.
- **Headings Detection:** Headings (H2-H4) are scanned periodically (every 2 seconds) or upon content updates to ensure the TOC remains in sync.

## Key Files
- `src/webparts/tableOfContents/TableOfContentsWebPart.ts`: Entry point and property pane definition.
- `src/webparts/tableOfContents/components/TableOfContents.tsx`: Orchestrator for TOC and Content.
- `src/webparts/tableOfContents/components/TipTapEditor.tsx`: Advanced editor logic and extensions.
- `src/webparts/tableOfContents/components/TOCNavigator.tsx`: TOC rendering logic.
- `package.json`: Dependency management and scripts.
- `config/`: SPFx-specific configuration files (manifests, package-solution, etc.).
