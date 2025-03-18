use pyo3::prelude::*;
mod book;
mod sheet;

/// A Python module implemented in Rust.
#[pymodule]
fn fastxlsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<book::BookWriter>()?;
    Ok(())
}
