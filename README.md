# ODoc

The ODoc markdown exporter; converting markdown files to pretty docx files.

## Usage

Coming soon.

## Showcase

Coming soon.

## Why?

### Formatting

I want to use markdown for productivity when creating essays which have to be exported to docx. I also want these exported docx files to look as professional and as well-presented as possible. The format is docx for two main reasons:

1. They're required for what I'm doing
2. They're very well-supported for format conversions (e.g. PDF, EPUB)

I'd rather directly export to something like PDF but it's less editable, convertible, and would be harder to develop for.

### Settings

As for settings, there are none. This project purposefully has no export settings as every valid markdown input should result in a heavily opinionated and *good looking* docx file.

### Language

Python was chosen as the language to develop this project with, in competition with Rust and Javascript. Python was chosen because:

- Rust doesn't have great docx bindings; it failed to compile on my machine
- Javascript is the worst of both worlds unless the situation requires clientside docx exporting on a website, which it doesn't
- Python has nice bindings and it makes it fast to hack together this project
