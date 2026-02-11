hello-world
===========

Demo of GitHub functionality

People of Earth:

The need for clear tracking of issues is upon us.  We must unite in our code repository keeping.

That is all.  More to come.

## SPFx Document Count Web Part

This repository now includes a lightweight SharePoint Framework (SPFx) web part that displays the number of documents in a selected library. The React component polls the library's `ItemCount` value on an interval so the number updates automatically without refreshing the page.

### Key files

- `src/webparts/documentCount/DocumentCountWebPart.ts`: SPFx web part class with property pane settings.
- `src/webparts/documentCount/components/DocumentCount.tsx`: React component that loads and refreshes the document count.
- `src/webparts/documentCount/components/DocumentCount.module.scss`: Styling for the rendered card.

### Usage notes

1. Set the **Library title** property to the document library you want to monitor (for example, `Documents`).
2. Adjust the **Refresh interval** slider to control how often the count is refreshed.
3. The component updates automatically in-place by calling the SharePoint REST API on each interval.
