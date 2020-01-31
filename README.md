# nb_extension_wordexport
Jupyter notebook extension for exporting notebook as MS Word doc

Install:

`pip install git+https://github.com/innovationOUtside/nb_extension_wordexport.git`

To upgrade a current installation to the latest repo version without updating dependencies:

`pip install --upgrade --no-deps git+https://github.com/innovationOUtside/nb_extension_wordexport.git`

## Wordexport

Requires `pandoc`: `!apt-get install -y pandoc`

Enable extension:

`jupyter bundlerextension enable --py wordexport.wordexport  --sys-prefix`

(If that doesn't work, try with `--user` rather than `--sys-prefix`.)

In TM351 VM, restart server if required:

`systemctl restart jupyter.service`


### Trying customisation:

//via http://stackoverflow.com/a/29484938/454773

`pandoc -D docx > my_template.docx`

Edit the styles in `my_template.docx`.

*I added a logo image in header but didn't seem to work in the following...*

`pandoc -s myfile.html --template=my_template -o test64.docx`


### Alternatives

This extension is quite old now... Others worth trying:

- [jupyter-docx-bundler](https://github.com/m-rossi/jupyter-docx-bundler)


See also:

- [nb2xls](https://github.com/ideonate/nb2xls): *convert Jupyter notebook to Excel spreadsheet*
