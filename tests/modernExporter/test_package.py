def test_package_importable():
    import deps.modernExporter as pkg
    assert pkg.__doc__ is not None
