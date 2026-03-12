use docxide_template::generate_templates;
use docxide_pdf::convert_docx_bytes_to_pdf;

generate_templates!("test-crate/templates");

fn main() {
    let hw = HelloWorld::new("World", "docxide");
    hw.save("test-crate/output/hello_world").unwrap();
    println!("Saved hello_world.docx");

    let table = TablePlaceholders::new("Alice", "Oslo");
    table.save("test-crate/output/table_placeholders").unwrap();
    println!("Saved table_placeholders.docx");

    let hf = HeadFootTest::new("My FooBarLOL", "droop", "Boom", "This goes down below");
    let hf_bytes = hf.to_bytes().unwrap();
    convert_docx_bytes_to_pdf(&hf_bytes, "test-crate/output/hello_world.pdf").unwrap();
    println!("Saved hello_world.pdf");

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

    // -- Formatted runs (bold/italic/underlined placeholders) --

    #[test]
    fn formatted_runs_struct_has_fields() {
        let f = FormattedRuns::new("Bold", "Mixed", "Underlined", "ABC");
        assert_eq!(f.bold_field, "Bold");
        assert_eq!(f.mixed_format, "Mixed");
        assert_eq!(f.underlined_field, "Underlined");
        assert_eq!(f.abc, "ABC");
    }

    #[test]
    fn formatted_runs_to_bytes_replaces() {
        let f = FormattedRuns::new("BoldVal", "MixedVal", "UnderVal", "ABCVal");
        let bytes = f.to_bytes().unwrap();
        let xml = read_zip_entry(&bytes, "word/document.xml");
        assert!(xml.contains("BoldVal"), "bold_field not replaced");
        assert!(xml.contains("MixedVal"), "mixed_format not replaced");
        assert!(xml.contains("UnderVal"), "underlined_field not replaced");
        assert!(xml.contains("ABCVal"), "abc not replaced");
        assert!(!xml.contains("{BoldField}"), "placeholder still present");
        assert!(!xml.contains("{ABC}"), "placeholder still present");
    }

    // -- Split runs template --

    #[test]
    fn split_runs_template_to_bytes_replaces() {
        let s = SplitRunsTemplate::new("Jane", "Acme Corp");
        let bytes = s.to_bytes().unwrap();
        let xml = read_zip_entry(&bytes, "word/document.xml");
        assert!(xml.contains("Jane"), "first_name not replaced");
        assert!(xml.contains("Acme Corp"), "company_name not replaced");
        assert!(!xml.contains("FirstName"), "placeholder still present");
        assert!(!xml.contains("CompanyName"), "placeholder still present");
    }

    // -- Unicode placeholders --

    #[test]
    fn unicode_placeholders_struct_has_fields() {
        let u = UnicodePlaceholders::new(
            "Ola", "Equinor", "Pilsner", "0.5L", "Jane", "ORD-123", "C-2024-1", "v2.1",
        );
        assert_eq!(u.fornavn, "Ola");
        assert_eq!(u.bedriftsnavn, "Equinor");
        assert_eq!(u.øltype, "Pilsner");
        assert_eq!(u.størrelse, "0.5L");
        assert_eq!(u.first_name, "Jane");
        assert_eq!(u.order_number, "ORD-123");
        assert_eq!(u.case2024_id, "C-2024-1");
        assert_eq!(u.app_version, "v2.1");
    }

    // -- Empty document --

    #[test]
    fn empty_document_generates_unit_struct() {
        use docxide_template::DocxTemplate;
        let e = EmptyDocument;
        assert!(e.replacements().is_empty());
    }

    // -- Cross-cutting concerns --

    #[test]
    fn xml_escaping_in_table_cells() {
        let t = TablePlaceholders::new("Alice & Bob", "<Oslo>");
        let bytes = t.to_bytes().unwrap();
        let xml = read_zip_entry(&bytes, "word/document.xml");
        assert!(xml.contains("Alice &amp; Bob"), "ampersand not escaped in table cell: {}", xml);
        assert!(xml.contains("&lt;Oslo&gt;"), "angle brackets not escaped in table cell: {}", xml);
        assert!(!xml.contains("Alice & Bob"), "raw ampersand should be escaped");
        assert!(!xml.contains("<Oslo>"), "raw angle brackets should be escaped");
    }

    #[test]
    fn xml_escaping_in_combined_areas() {
        let c = CombinedAreas::new("A & B", "x < y", "1 > 0", "R&D \"Report\"", "page'7");
        let bytes = c.to_bytes().unwrap();
        let all = all_xml_content(&bytes);
        assert!(all.contains("A &amp; B"), "ampersand not escaped in body");
        assert!(all.contains("x &lt; y"), "less-than not escaped in table");
        assert!(all.contains("1 &gt; 0"), "greater-than not escaped in table");
        assert!(all.contains("R&amp;D &quot;Report&quot;"), "quote not escaped in header");
        assert!(all.contains("page&apos;7"), "apostrophe not escaped in footer");
    }

    #[test]
    fn to_bytes_output_is_valid_zip_for_all_templates() {
        let templates: Vec<Vec<u8>> = vec![
            HelloWorld::new("A", "B").to_bytes().unwrap(),
            TablePlaceholders::new("A", "B").to_bytes().unwrap(),
            HeadFootTest::new("A", "B", "C", "D").to_bytes().unwrap(),
            CombinedAreas::new("A", "B", "C", "D", "E").to_bytes().unwrap(),
            TextboxTemplate::new("A", "B", "C").to_bytes().unwrap(),
            FormattedRuns::new("A", "B", "C", "D").to_bytes().unwrap(),
            SplitRunsTemplate::new("A", "B").to_bytes().unwrap(),
        ];
        for (i, bytes) in templates.iter().enumerate() {
            assert!(!bytes.is_empty(), "template {} produced empty output", i);
            let cursor = Cursor::new(bytes);
            let archive = zip::ZipArchive::new(cursor);
            assert!(archive.is_ok(), "template {} produced invalid zip: {:?}", i, archive.err());
            assert!(archive.unwrap().len() > 0, "template {} zip has no entries", i);
        }
    }

    #[test]
    fn save_and_to_bytes_produce_same_xml() {
        let hw = HelloWorld::new("SaveTest", "BytesTest");
        let bytes_output = hw.to_bytes().unwrap();

        let tmp_dir = std::env::temp_dir().join("docxide_test_save_vs_bytes");
        let _ = std::fs::create_dir_all(&tmp_dir);
        let save_path = tmp_dir.join("compare");
        hw.save(&save_path).unwrap();

        let saved_bytes = std::fs::read(save_path.with_extension("docx")).unwrap();
        let _ = std::fs::remove_dir_all(&tmp_dir);

        let to_bytes_xml = all_xml_content(&bytes_output);
        let saved_xml = all_xml_content(&saved_bytes);
        assert_eq!(to_bytes_xml, saved_xml, "save() and to_bytes() should produce identical XML");
    }
}

