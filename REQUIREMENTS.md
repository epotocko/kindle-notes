# Kindle Notes Viewer

The Kindle Notes Viewer is a web application that allows users to view their Kindle notes and highlights. It supports modern web browsers and the Kindle web browser. The initial view will be a calendar view that shows the number of notes/highlights for each day of the month. Users can click on a specific day to view the notes for that day. 

## Requirements

* SPA built using Vanilla JavaScript (no frameworks)
  * No backend
  * Requirements can be added such as MSAL.js for auth
  * JavaScript libraries can be added for bundling the app
  * All Javascript must be supported by the Kindle scribe
* Hosted on GitHub Pages
  * Github Actions for publishing
* Desktop view or Kindle view (black-and-white, eInk friendly) - can we do this with themes via css?
* Home page provides options to retrieve notes
  * Retrieve notes from local file (e.g., `My Clippings.txt`)
  * Retrieve notes from OneDrive (requires user authentication)
* Main page has tabs for: Month, Day
  * Day displays all notes from the current day and allows users to navigate to previous/next days
  * Month displays the number of notes/highlights for each day of the month. Users can click on a specific day to view the notes for that day.
