# Purpose

This module acts as a custom documentation generator for Stencil projects, and produces a Word .docx document
containing a listing of components and their `jsdoc` documentation strings for props and slots.

# Usage

Add an `outputTarget` to your `stencil.config.ts`:

```
import { Config } from '@stencil/core';
import createDocxGenerator from "stencil-docx-docs";

// https://stenciljs.com/docs/config

export const config: Config = {
  globalStyle: 'src/global/app.css',
  globalScript: 'src/global/app.ts',
  taskQueue: 'async',
  outputTargets: [
    {
      type: 'www',
      // comment the following line to disable service workers in production
      serviceWorker: null,
      baseUrl: 'https://myapp.local/',
    },
    {
      type: "docs-custom",
      generator: createDocxGenerator({
        // options - see below
      }),
    }
  ],
};
```

The following options are available:

| Option        | Default                   | Description                                                            |
| ------------- | ------------------------- | ---------------------------------------------------------------------- |
| `outDir`      | `docs`                    | The output directory                                                   |
| `outFile`     | `docs.docx`               | The output file name                                                   |
| `textFont`    | `Calibri`                 | The font that will be used for all text in the document                |
| `excludeTags` | `["undocumented"]`        | An array of doc tags that will cause a component or prop to be ignored |
| `title`       | `Component Documentation` | A document title                                                       |
| `author`      | `stencil-docx-docs`       | The author of the document                                             |
