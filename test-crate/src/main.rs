use docxide_template::generate_templates;

generate_templates!("test-crate/templates");

fn main() {
    let hw = HelloWorld::new("World", "docxide");
    hw.save("test-crate/output/hello_world").unwrap();
    println!("Saved hello_world.docx");

    let table = TablePlaceholders::new("Alice", "Oslo");
    table.save("test-crate/output/table_placeholders").unwrap();
    println!("Saved table_placeholders.docx");

    let hf = HeadFootTest::new("My Report", "42", "This goes on top", "This goes down below");
    hf.save("test-crate/output/header_footer_placeholders").unwrap();
    println!("Saved header_footer_placeholders.docx");

    let combined = CombinedAreas::new("Bob", "Item", "100", "Quarterly Report", "7");
    combined.save("test-crate/output/combined_areas").unwrap();
    println!("Saved combined_areas.docx");
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Read};

    fn read_zip_entry(docx_bytes: &[u8], entry_name: &str) -> String {
        let cursor = Cursor::new(docx_bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut content = String::new();
        archive
            .by_name(entry_name)
            .unwrap()
            .read_to_string(&mut content)
            .unwrap();
        content
    }

    fn all_xml_content(docx_bytes: &[u8]) -> String {
        let cursor = Cursor::new(docx_bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut combined = String::new();
        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            if name.ends_with(".xml") || name.ends_with(".rels") {
                let mut content = String::new();
                file.read_to_string(&mut content).unwrap();
                combined.push_str(&content);
            }
        }
        combined
    }

    // -- Table placeholders --

    #[test]
    fn table_placeholders_struct_has_fields() {
        let t = TablePlaceholders::new("Alice", "Oslo");
        assert_eq!(t.table_name, "Alice");
        assert_eq!(t.table_city, "Oslo");
    }

    #[test]
    fn table_placeholders_to_bytes_replaces() {
        let t = TablePlaceholders::new("Alice", "Oslo");
        let bytes = t.to_bytes().unwrap();
        let xml = read_zip_entry(&bytes, "word/document.xml");
        assert!(xml.contains("Alice"), "table_name not replaced");
        assert!(xml.contains("Oslo"), "table_city not replaced");
        assert!(!xml.contains("table_name"), "placeholder still present");
        assert!(!xml.contains("table_city"), "placeholder still present");
    }

    // -- Header/footer placeholders --

    #[test]
    fn header_footer_struct_has_fields() {
        // Template has: body {header} {foo}, header {top}, footer {bottom}
        let hf = HeadFootTest::new("My Report", "42", "Banner", "Fine Print");
        assert_eq!(hf.header, "My Report");
        assert_eq!(hf.foo, "42");
        assert_eq!(hf.top, "Banner");
        assert_eq!(hf.bottom, "Fine Print");
    }

    #[test]
    fn header_footer_to_bytes_replaces() {
        let hf = HeadFootTest::new("My Report", "42", "Banner", "Fine Print");
        let bytes = hf.to_bytes().unwrap();
        let all = all_xml_content(&bytes);
        assert!(all.contains("Banner"), "header top not replaced");
        assert!(all.contains("Fine Print"), "footer bottom not replaced");
    }

    #[test]
    fn header_footer_replacements_include_whitespace_variants() {
        use docxide_template::DocxTemplate;
        let hf = HeadFootTest::new("My Report", "42", "Banner", "Fine Print");
        let reps = hf.replacements();
        let placeholders: Vec<&str> = reps.iter().map(|(p, _)| *p).collect();
        assert!(
            placeholders.contains(&"{ foo }"),
            "missing {{ foo }} in replacements: {:?}",
            placeholders,
        );
        assert!(
            placeholders.contains(&"{  foo  }"),
            "missing {{  foo  }} in replacements: {:?}",
            placeholders,
        );
    }

    // -- Combined areas --

    #[test]
    fn combined_areas_struct_has_all_fields() {
        // Field order: body paragraph, table cells, header, footer
        let c = CombinedAreas::new("Bob", "Item", "100", "Report", "7");
        assert_eq!(c.body_name, "Bob");
        assert_eq!(c.cell_label, "Item");
        assert_eq!(c.cell_value, "100");
        assert_eq!(c.doc_title, "Report");
        assert_eq!(c.page_num, "7");
    }

    // -- Text box placeholders --

    #[test]
    fn textbox_template_struct_has_textbox_field() {
        let t = TextboxTemplate::new("Alice", "Widget", "Boxed");
        assert_eq!(t.first_name, "Alice");
        assert_eq!(t.product_name, "Widget");
        assert_eq!(t.textbox_field, "Boxed");
    }

    #[test]
    fn textbox_template_to_bytes_replaces() {
        let t = TextboxTemplate::new("Alice", "Widget", "Boxed");
        let bytes = t.to_bytes().unwrap();
        let xml = read_zip_entry(&bytes, "word/document.xml");
        assert!(xml.contains("Boxed"), "textbox_field not replaced");
        assert!(!xml.contains("textbox_field"), "placeholder still present");
    }

    // -- Combined areas --

    #[test]
    fn combined_areas_to_bytes_replaces_all() {
        let c = CombinedAreas::new("Bob", "Item", "100", "Report", "7");
        let bytes = c.to_bytes().unwrap();
        let all = all_xml_content(&bytes);
        assert!(all.contains("Report"), "doc_title not replaced");
        assert!(all.contains("Bob"), "body_name not replaced");
        assert!(all.contains("Item"), "cell_label not replaced");
        assert!(all.contains("100"), "cell_value not replaced");
        assert!(!all.contains("doc_title"), "placeholder still present");
        assert!(!all.contains("body_name"), "placeholder still present");
        assert!(!all.contains("cell_label"), "placeholder still present");
        assert!(!all.contains("cell_value"), "placeholder still present");
    }
}

