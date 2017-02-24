officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->
[![Build Status](https://travis-ci.org/davidgohel/officer.svg?branch=master)](https://travis-ci.org/davidgohel/officer) [![Coverage Status](https://img.shields.io/codecov/c/github/davidgohel/officer/master.svg)](https://codecov.io/github/davidgohel/officer?branch=master) [![CRAN version](http://www.r-pkg.org/badges/version/officer)](http://cran.r-project.org/package=officer) ![](http://cranlogs.r-pkg.org/badges/grand-total/officer) [![Project Status: WIP - Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](http://www.repostatus.org/badges/latest/wip.svg)](http://www.repostatus.org/#wip)

The officer package lets R users manipulate Word (`.docx`) and PowerPoint (`*.pptx`) documents. In short, one can add images, tables and text into documents from R.

*This package is close to ReporteRs as it produces Word and PowerPoint files but it is faster, do not require `rJava` (but `xml2`) and has less functions that will make it easier to maintain.*

------------------------------------------------------------------------

<img src="data:image/png;base64,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" width="128" height="128"/>

Function `read_docx` will read an initial Word document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into a Word document:

-   images: produce your plot in png or emf files and add them into the document; as a whole paragraph or inside a paragraph.
-   tables: add data.frames as tables, format is defined by the associated Word table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined by the associated Word paragraph and text styles.
-   field codes: add Word field codes inside paragraphs. Field codes is an old feature of MS Word to create calculated elements such as tables of content, automatic numbering and hyperlinks.

In a Word document, one can use cursor functions to reach the beginning of a document, its end or a particular paragraph containing a given text. This *cursor* concept has been implemented to make easier the post processing of files.

The file generation is performed with function `print`.

------------------------------------------------------------------------

<img src="data:image/png;base64,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" width="128" height="128"/>

Function `read_pptx` will read an initial PowerPoint document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into existing or new PowerPoint slides:

-   images: produce your plot in png or emf files and add them in a slide.
-   tables: add data.frames as tables, format is defined by the associated PowerPoint table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined in the corresponding layout of the slide.

In a PowerPoint document, one can set a slide as selected and reach a particular shape (and remove it or add text).

The file generation is performed with function `print`.

### Tables and package `flextable`

The package [flextable](https://github.com/davidgohel/flextable) brings a full API to produce nice tables and use them with `officer`.

### Installation

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
```

Or the latest version on CRAN:

``` r
install.packages("officer")
```
