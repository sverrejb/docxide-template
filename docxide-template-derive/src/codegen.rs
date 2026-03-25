use quote::quote;

pub(crate) fn generate_struct(
    type_ident: syn::Ident,
    abs_path: &str,
    rel_path: &str,
    fields: &[syn::Ident],
    replacement_placeholders: &[syn::LitStr],
    replacement_fields: &[syn::Ident],
    embed: bool,
) -> proc_macro2::TokenStream {
    let has_fields = !fields.is_empty();
    let abs_path_lit = syn::LitStr::new(abs_path, proc_macro::Span::call_site().into());
    let rel_path_lit = syn::LitStr::new(rel_path, proc_macro::Span::call_site().into());

    let to_bytes_impl = if embed {
        quote! {
            fn to_bytes(&self) -> Result<Vec<u8>, docxide_template::TemplateError> {
                docxide_template::__private::build_docx_bytes(Self::TEMPLATE_BYTES, &self.replacements())
            }
        }
    } else {
        quote! {
            fn to_bytes(&self) -> Result<Vec<u8>, docxide_template::TemplateError> {
                let template_bytes = std::fs::read(self.template_path())?;
                docxide_template::__private::build_docx_bytes(&template_bytes, &self.replacements())
            }
        }
    };

    let embed_const = if embed {
        quote! { const TEMPLATE_BYTES: &'static [u8] = include_bytes!(#abs_path_lit); }
    } else {
        quote! {}
    };

    if has_fields {
        quote! {
            #[derive(Debug, Clone)]
            pub struct #type_ident {
                #(pub #fields: String,)*
            }

            impl #type_ident {
                #embed_const

                pub fn new(#(#fields: impl Into<String>),*) -> Self {
                    Self {
                        #(#fields: #fields.into()),*
                    }
                }

                pub fn to_bytes(&self) -> Result<Vec<u8>, docxide_template::TemplateError> {
                    <Self as docxide_template::DocxTemplate>::to_bytes(self)
                }

                pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<(), docxide_template::TemplateError> {
                    <Self as docxide_template::DocxTemplate>::save(self, path.as_ref().with_extension("docx").as_path())
                }
            }

            impl docxide_template::DocxTemplate for #type_ident {
                fn template_path(&self) -> &std::path::Path {
                    std::path::Path::new(#rel_path_lit)
                }

                fn replacements(&self) -> Vec<(&str, &str)> {
                    vec![#( (#replacement_placeholders, self.#replacement_fields.as_str()), )*]
                }

                #to_bytes_impl
            }
        }
    } else {
        quote! {
            #[derive(Debug, Clone)]
            pub struct #type_ident;

            impl #type_ident {
                #embed_const

                pub fn to_bytes(&self) -> Result<Vec<u8>, docxide_template::TemplateError> {
                    <Self as docxide_template::DocxTemplate>::to_bytes(self)
                }

                pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<(), docxide_template::TemplateError> {
                    <Self as docxide_template::DocxTemplate>::save(self, path.as_ref().with_extension("docx").as_path())
                }
            }

            impl docxide_template::DocxTemplate for #type_ident {
                fn template_path(&self) -> &std::path::Path {
                    std::path::Path::new(#rel_path_lit)
                }

                fn replacements(&self) -> Vec<(&str, &str)> {
                    vec![]
                }

                #to_bytes_impl
            }
        }
    }
}
