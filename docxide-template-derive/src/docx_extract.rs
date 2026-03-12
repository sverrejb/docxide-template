use docx_rs::{
    DocumentChild, DrawingData, FooterChild, HeaderChild, Paragraph, ParagraphChild, RunChild,
    StructuredDataTagChild, Table, TableCellContent, TableChild, TableRowChild,
    TextBoxContentChild,
};
use file_format::FileFormat;
use std::io::{Cursor, Read};
use std::path::Path;

pub(crate) fn collect_text_from_document_children(children: Vec<DocumentChild>) -> Vec<String> {
    let mut texts = Vec::new();
    for child in children {
        match child {
            DocumentChild::Paragraph(p) => texts.extend(collect_text_from_paragraph(&p)),
            DocumentChild::Table(t) => texts.extend(collect_text_from_table(&t)),
            DocumentChild::StructuredDataTag(sdt) => {
                texts.extend(collect_text_from_sdt_children(&sdt.children));
            }
            _ => {}
        }
    }
    texts
}

pub(crate) fn collect_text_from_table(table: &Table) -> Vec<String> {
    let mut texts = Vec::new();
    for row in &table.rows {
        let TableChild::TableRow(ref row) = row;
        for cell in &row.cells {
            let TableRowChild::TableCell(ref cell) = cell;
            for content in &cell.children {
                match content {
                    TableCellContent::Paragraph(p) => texts.extend(collect_text_from_paragraph(p)),
                    TableCellContent::Table(t) => texts.extend(collect_text_from_table(t)),
                    _ => {}
                }
            }
        }
    }
    texts
}

fn collect_text_from_sdt_children(children: &[StructuredDataTagChild]) -> Vec<String> {
    let mut texts = Vec::new();
    for child in children {
        match child {
            StructuredDataTagChild::Paragraph(p) => texts.extend(collect_text_from_paragraph(p)),
            StructuredDataTagChild::Table(t) => texts.extend(collect_text_from_table(t)),
            StructuredDataTagChild::StructuredDataTag(sdt) => {
                texts.extend(collect_text_from_sdt_children(&sdt.children));
            }
            _ => {}
        }
    }
    texts
}

pub(crate) fn collect_text_from_header_children(children: &[HeaderChild]) -> Vec<String> {
    let mut texts = Vec::new();
    for child in children {
        match child {
            HeaderChild::Paragraph(p) => texts.extend(collect_text_from_paragraph(p)),
            HeaderChild::Table(t) => texts.extend(collect_text_from_table(t)),
            HeaderChild::StructuredDataTag(sdt) => {
                texts.extend(collect_text_from_sdt_children(&sdt.children));
            }
        }
    }
    texts
}

pub(crate) fn collect_text_from_footer_children(children: &[FooterChild]) -> Vec<String> {
    let mut texts = Vec::new();
    for child in children {
        match child {
            FooterChild::Paragraph(p) => texts.extend(collect_text_from_paragraph(p)),
            FooterChild::Table(t) => texts.extend(collect_text_from_table(t)),
            FooterChild::StructuredDataTag(sdt) => {
                texts.extend(collect_text_from_sdt_children(&sdt.children));
            }
        }
    }
    texts
}

fn collect_text_from_paragraph(p: &Paragraph) -> Vec<String> {
    let mut texts = vec![p.raw_text()];
    for child in &p.children {
        let runs: Vec<&RunChild> = match child {
            ParagraphChild::Run(run) => run.children.iter().collect(),
            ParagraphChild::Insert(ins) => ins
                .children
                .iter()
                .filter_map(|c| {
                    if let docx_rs::InsertChild::Run(r) = c {
                        Some(r.children.iter())
                    } else {
                        None
                    }
                })
                .flatten()
                .collect(),
            ParagraphChild::Hyperlink(h) => h
                .children
                .iter()
                .filter_map(|c| {
                    if let ParagraphChild::Run(r) = c {
                        Some(r.children.iter())
                    } else {
                        None
                    }
                })
                .flatten()
                .collect(),
            _ => continue,
        };
        for run_child in runs {
            if let RunChild::Drawing(drawing) = run_child {
                if let Some(DrawingData::TextBox(text_box)) = &drawing.data {
                    texts.extend(collect_text_from_textbox_content(&text_box.children));
                }
            }
        }
    }
    texts
}

fn collect_text_from_textbox_content(children: &[TextBoxContentChild]) -> Vec<String> {
    let mut texts = Vec::new();
    for child in children {
        match child {
            TextBoxContentChild::Paragraph(p) => texts.extend(collect_text_from_paragraph(p)),
            TextBoxContentChild::Table(t) => texts.extend(collect_text_from_table(t)),
        }
    }
    texts
}

/// Extracts text content from multiple XML tag types in a single zip pass.
/// Each `open_prefix` (e.g. `"<a:t"`) is paired with a close tag derived
/// automatically (e.g. `"</a:t>"`).
pub(crate) fn extract_text_from_xml_tags(docx_bytes: &[u8], open_prefixes: &[&str]) -> Vec<String> {
    let tags: Vec<(&str, String)> = open_prefixes
        .iter()
        .map(|p| (*p, format!("</{}", &p[1..])))
        .collect();

    let mut texts = Vec::new();
    let cursor = Cursor::new(docx_bytes);
    let mut archive = match zip::ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(_) => return texts,
    };

    for i in 0..archive.len() {
        let mut file = match archive.by_index(i) {
            Ok(f) => f,
            Err(_) => continue,
        };
        if !file.name().ends_with(".xml") {
            continue;
        }
        let mut xml = String::new();
        if file.read_to_string(&mut xml).is_err() {
            continue;
        }

        for (open_prefix, close_tag) in &tags {
            let mut search_start = 0;
            while let Some(tag_start) = xml[search_start..].find(open_prefix) {
                let tag_start = search_start + tag_start;
                let after_prefix = tag_start + open_prefix.len();
                if after_prefix < xml.len() && !matches!(xml.as_bytes()[after_prefix], b'>' | b' ') {
                    search_start = after_prefix;
                    continue;
                }
                let content_start = match xml[tag_start..].find('>') {
                    Some(pos) => tag_start + pos + 1,
                    None => break,
                };
                let content_end = match xml[content_start..].find(close_tag.as_str()) {
                    Some(pos) => content_start + pos,
                    None => break,
                };
                let text = xml[content_start..content_end].to_string();
                if !text.is_empty() {
                    texts.push(text);
                }
                search_start = content_end + close_tag.len();
            }
        }
    }

    texts
}

pub(crate) fn print_docxide_message(message: &str, path: &Path) {
    println!("\x1b[34m[Docxide-template]\x1b[0m {} {:?}", message, path);
}

pub(crate) fn is_valid_docx_file(path: &Path) -> bool {
    if !path.is_file() {
        return false;
    }

    matches!(FileFormat::from_file(path), Ok(fmt) if fmt.extension() == "docx")
}
