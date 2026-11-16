use rust_office_art::{PictureContext, ResolvedColor, StyleReferenceKind};
use wasm_bindgen::prelude::*;

struct Context {
    colors: js_sys::Function,
    styles: Option<js_sys::Function>,
}

#[derive(serde::Deserialize)]
struct Color {
    color: String,
    opacity: Option<f64>,
}

impl PictureContext for Context {
    fn resolve_color(&self, xml: &str) -> Option<ResolvedColor> {
        let value = self
            .colors
            .call1(&JsValue::NULL, &JsValue::from_str(xml))
            .ok()?
            .as_string()?;
        let color: Color = serde_json::from_str(&value).ok()?;
        Some(ResolvedColor {
            color: color.color,
            opacity: color.opacity,
        })
    }

    fn resolve_style_reference(
        &self,
        kind: StyleReferenceKind,
        index: u32,
        placeholder_xml: &str,
    ) -> Option<String> {
        let kind = match kind {
            StyleReferenceKind::Fill => "fill",
            StyleReferenceKind::Line => "line",
            StyleReferenceKind::Effect => "effect",
        };
        self.styles
            .as_ref()?
            .call3(
                &JsValue::NULL,
                &JsValue::from_str(kind),
                &JsValue::from(index),
                &JsValue::from_str(placeholder_xml),
            )
            .ok()?
            .as_string()
    }
}

#[wasm_bindgen(js_name = renderPicture)]
pub fn render_picture(
    xml: &str,
    source: &str,
    width: f64,
    height: f64,
    colors: js_sys::Function,
) -> Result<Option<String>, JsError> {
    let context = Context {
        colors,
        styles: None,
    };
    rust_office_art::render_picture(xml, source, width, height, &context)
        .map(|picture| serde_json::to_string(&picture).map_err(JsError::from))
        .transpose()
}

/// Render a shape or connector's geometry, fill and outline. `styles` resolves
/// a theme style reference: it receives "fill", "line" or "effect", the
/// reference's index and its XML, and returns the theme style's XML with the
/// placeholder colour applied, or undefined.
#[wasm_bindgen(js_name = renderShape)]
pub fn render_shape(
    properties_xml: &str,
    style_xml: &str,
    width: f64,
    height: f64,
    colors: js_sys::Function,
    styles: js_sys::Function,
) -> Result<Option<String>, JsError> {
    let context = Context {
        colors,
        styles: Some(styles),
    };
    rust_office_art::render_shape(properties_xml, style_xml, width, height, &context)
        .map(|shape| serde_json::to_string(&shape).map_err(JsError::from))
        .transpose()
}

#[wasm_bindgen(js_name = normalizeRaster)]
pub fn normalize_raster(source: &str) -> Option<String> {
    rust_office_art::normalize_raster(source)
}
