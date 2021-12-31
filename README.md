# DocMerge

A mail merge utility for use with Google Docs and Google Sheets implemented in Apps Script.

## Usage

The required inputs are the ID of the template document, the starting row, and ending row (inclusive). The ID of the template document can be found by opening the Google Doc in your browser - the URL should be `https://docs.google.com/document/d/<ID>/edit`.

### Spreadsheet

The first row of the spreadsheet defines the key for the column. Each following row represents a set of entries to populate the template document.

### Template Document

The template document should contain positions for key replacements. The keys are marked using the pattern `${<key>}`. For example, a key `First Name` would be marked using `${First Name}`.
