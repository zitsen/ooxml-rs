#[test]
fn test_zip_writer() -> Result<(), crate::error::OoxmlError>{
    use std::io::Write;

    // We use a buffer here, though you'd normally use a `File`
    //let mut buf = [0; 65536];
    //let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut buf[..]));
    let file = std::fs::File::create("tests/test.zip")?;
    let mut zip = zip::ZipWriter::new(file);

    let options =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("test/hello_world.txt", options)?;
    zip.write(b"Hello, World!")?;

    // Apply the changes you've made.
    // Dropping the `ZipWriter` will have the same effect, but may silently fail
    zip.finish()?;
    Ok(())
}
