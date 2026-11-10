def test_package_importable():
    import deps.exporter as pkg

    assert pkg.__doc__ is not None
