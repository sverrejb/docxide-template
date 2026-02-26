use docx_rs::{
    DocumentChild, DrawingData, FooterChild, HeaderChild, Paragraph, ParagraphChild, RunChild,
    StructuredDataTagChild, Table, TableCellContent, TableChild, TableRowChild,
    TextBoxContentChild,
};
use file_format::FileFormat;
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

pub(crate) fn print_docxide_message(message: &str, path: &Path) {
    println!("\x1b[34m[Docxide-template]\x1b[0m {} {:?}", message, path);
}

pub(crate) fn is_valid_docx_file(path: &Path) -> bool {
    if !path.is_file() {
        return false;
    }

    matches!(FileFormat::from_file(path), Ok(fmt) if fmt.extension() == "docx")
}
