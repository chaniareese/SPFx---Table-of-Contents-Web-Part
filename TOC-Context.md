[PROJECT_IDENTITY]
Name: SPFx Table of Contents Web Part
Type: SharePoint Framework (SPFx) Client-Side Web Part
Version: 0.0.1
SPFx Version: 1.22.2

[CONTEXT_OVERVIEW]
This project is a specialized SharePoint web part designed to provide a "Runbook" or "Documentation" experience. It features a dual-pane layout: a dynamic, auto-generated Table of Contents (TOC) on the left and a sophisticated rich-text editor (TipTap) on the right.

[TECHNICAL_STACK]
- UI Framework: React 17.0.1
- Core Library: @microsoft/sp-webpart-base
- Editor Engine: TipTap 3.20.1 (Headless Rich-Text)
- Styling: SASS/SCSS Modules
- Build Toolchain: @rushstack/heft
- Package Manager: npm (requires Node.js 22.x)

[ARCHITECTURAL_COMPONENTS]
1. TableOfContentsWebPart.ts:
   - Extends BaseClientSideWebPart.
   - Manages Property Pane (toggle H2/H3/H4, mobile visibility, title).
   - Handles property persistence.

2. TableOfContents.tsx:
   - Primary state orchestrator.
   - Implements heading extraction logic (querySelectorAll('h2, h3, h4')).
   - Manages scroll-sync to highlight active sections in the sidebar.

3. TipTapEditor.tsx:
   - Custom implementation of TipTap.
   - Features: Tables (with custom styling), Resizable Images, Font Size, Line Height, Indentation.
   - Uses Module Augmentation for custom TypeScript command definitions.

4. TOCNavigator.tsx:
   - Functional component for the sidebar.
   - Recursively builds a nested <ul> structure based on heading levels.
   - Provides smooth-scroll navigation.

[LIFECYCLE_COMMANDS]
- Local Dev: heft start
- Clean: heft clean
- Build Production: heft test --clean --production
- Package: heft package-solution --production

[DEVELOPMENT_CONVENTIONS]
- Use SCSS modules for all component styling.
- Maintain strict TypeScript types.
- Update the heading scanner if new block types are added to the editor.
- Sync editor content back to web part properties via the updateProperty callback.

[FILE_MAP]
- /src/webparts/tableOfContents/TableOfContentsWebPart.ts (Entry)
- /src/webparts/tableOfContents/components/TableOfContents.tsx (Main)
- /src/webparts/tableOfContents/components/TipTapEditor.tsx (Editor)
- /src/webparts/tableOfContents/components/TOCNavigator.tsx (Sidebar)
- /config/package-solution.json (Manifest)
- /package.json (Dependencies)
