PDF Word Counter
==============

![Last tested in Nuix 7.6](https://img.shields.io/badge/Nuix-7.6-green.svg)

View the GitHub project [here](https://github.com/Nuix/PDF-Word-Counter) or download the latest release [here](https://github.com/Nuix/PDF-Word-Counter/releases).

View the Docs [here](https://nuix.github.io/PDF-Word-Counter/).

# Overview

**Written By:** mrk

This script adds custom metadata fields for "Word Count" and "Words Per Page" to searchable PDFs.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

## Usage

Before running the script you will need to have a case open in Nuix.

The script counts the number of words in an item's text, and creates custom metadata fields for "Word Count" and "Words Per Page".

### Options

The script's options are in count_words.rb and include:
- `query:` the search query for the items to process. (Default is searchable PDFs, i.e. `mime-type:application/pdf AND content:*'`)
- `handleExcluded:` if exclusions should be considered. (i.e. `AND has-exclusion:0`)
- `recalculate:` if custom metadata values should be recalculated for items that already have the field.

### Word Count

The custom metadata for "Word Count" is calculated by taking the item's text string, converting non-alphabet characters to space, and splitting on whitespace (i.e. `String.gsub(/[^-a-zA-Z]/, ' ').split.size`). This works because `string.split` splits on whitespace with leading whitespace and runs of contiguous whitespace characters ignored.

Search Examples:
- Less than 20 words: `integer-custom-metadata-exact:"Word Count":[0 TO 20}`
- Between 20 and 40 words: `integer-custom-metadata-exact:"Word Count":[20 TO 40]`
- Over 100 words: `integer-custom-metadata-exact:"Word Count":{100 TO *]`

### Words Per Page

The custom metadata for "Words Per Page" is a float, dividing the word count by the number of pages. The number of pages is determined by the metadata property "Page Count" (or 1 if the property is nil).

Search Examples:
- Less than 1 word per page: `float-custom-metadata-exact:"Words Per Page":[0 TO 1}`
- At least 5.5 words per page: `float-custom-metadata-exact:"Words Per Page":[5.5 TO *]`

## Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the scripts.

# License

```
Copyright 2018 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
