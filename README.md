# nb_extension_wordexport
Jupyter notebook extension for exporting notebook as MS Word doc


---

Recent versions of `pandoc` now include native support for `.ipynb`, which means we can have a direct export of .`ipynb` to `.docx` via pandoc.


Install upgraded pandoc:

```
wget https://github.com/jgm/pandoc/releases/download/2.7.3/pandoc-2.7.3-linux.tar.gz
sudo tar xvzf pandoc-2.7.3-linux.tar.gz --strip-components 1 -C /usr/local
```



---

Install:

`pip install git+https://github.com/innovationOUtside/nb_extension_wordexport.git`

To upgrade a current installation to the latest repo version without updating dependencies:

`pip install --upgrade --no-deps git+https://github.com/innovationOUtside/nb_extension_wordexport.git`

## Wordexport

Requires `pandoc`: `!apt-get install -y pandoc`

Enable extension:

`jupyter bundlerextension enable --py wordexport.wordexport  --sys-prefix`

In TM351 VM, restart server if required:

`systemctl restart jupyter.service`


### Trying customisation:

//via http://stackoverflow.com/a/29484938/454773

`pandoc -D docx > my_template.docx`

Edit the styles in `my_template.docx`.

*I added a logo image in header but didn't seem to work in the following...*

`pandoc -s myfile.html --template=my_template -o test64.docx`

