# Configuration file for the Sphinx documentation builder.

# -- Project information

project = 'LibreOffice Basic Reference'
copyright = '2024, Ray Culp'
author = 'Ray Culp'

release = '0.1'
version = '0.1.0'

# -- General configuration

extensions = [
    'sphinx.ext.duration',
    'sphinx.ext.doctest',
    'sphinx.ext.autodoc',
    'sphinx.ext.autosummary',
    'sphinx.ext.intersphinx',
]

intersphinx_mapping = {
    'python': ('https://docs.python.org/3/', None),
    'sphinx': ('https://www.sphinx-doc.org/en/master/', None),
}
intersphinx_disabled_domains = ['std']

templates_path = ['_templates']

html_static_path = ['_static']

# -- Options for HTML output

# html_theme = 'furo'
html_theme = 'sphinx_rtd_theme'
html_theme_options = {
    'collapse_navigation': False,  # Keeps the navigation expanded
    'navigation_depth': 10,         # Adjust this based on your TOC depth
    'titles_only': True,          # Show all titles in the TOC
}

extensions = [
    'sphinx_copybutton'
]

html_css_files = ['css/custom.css']

# -- Options for EPUB output
epub_show_urls = 'footnote'


