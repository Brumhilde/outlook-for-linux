# Custom CSS Module

Injects a user-supplied stylesheet into the Outlook page after each load.

## Configuration Options

- **`appearance.cssLocation`**: Path to a custom CSS file

Outlook renders in the top-level frame, so the CSS is inserted with `webContents.insertCSS()`. Outlook's class names are generated and change between releases, so prefer attribute selectors (`[role=...]`, `[aria-label=...]`) in custom styles.

Configuration details: [`../../docs/configuration.md`](../../docs/configuration.md)
