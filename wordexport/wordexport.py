# Copyright (c) The Open University, 2017
# Copyright (c) Jupyter Development Team.

# Distributed under the terms of the Modified BSD License.

# Custom bundler extensions API:
# http://jupyter-notebook.readthedocs.io/en/latest/extending/bundler_extensions.html
# Requires: _jupyter_bundlerextension_paths(), bundle()
# Based on: https://github.com/jupyter-incubator/dashboards_bundlers/

import os
import subprocess
import shlex
import shutil
import tempfile

def _jupyter_bundlerextension_paths():
    '''API for notebook bundler installation on notebook 5.0+'''
    return [{
                'name': 'wordexport_wordexport',
                'label': 'MS Word (.docx)',
                'module_name': 'wordexport.wordexport',
                'group': 'download'
            }]

def bundle(handler, model):
	'''
	Downloads a notebook as a Microsoft Word .docx document after converting to HTML
	'''
	
	# Based on https://github.com/jupyter-incubator/dashboards_bundlers
	
	abs_nb_path = os.path.join(
		handler.settings['contents_manager'].root_dir,
		model['path']
	)
		
	notebook_basename = os.path.basename(abs_nb_path)
	notebook_name = os.path.splitext(notebook_basename)[0]
	
	tmp_dir = tempfile.mkdtemp()
	
	# Generate HTML version of file with embedded images using --to html_embed 
	#  causes pandoc error 
	cmd='jupyter nbconvert --to html "{abs_nb_path}" --output-dir "{tmp_dir}"'.format(abs_nb_path=abs_nb_path,tmp_dir=tmp_dir)
	#os.system(cmd)
	subprocess.check_call(cmd, shell=True)
	
	staged=os.path.join(tmp_dir, notebook_name)
	# Convert to MS Word .docx
	cmd='pandoc -s "{staged}.html" -o "{staged}.docx"'.format(staged=staged)
	#os.system(cmd)
	subprocess.check_call(cmd, shell=True)
		
	handler.set_header('Content-Disposition', 'attachment; filename="%s"' % (notebook_name + '.docx'))
	
	handler.set_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
	
	with open("{staged}.docx".format(staged=staged), 'rb') as bundle_file:
		handler.write(bundle_file.read())

	handler.finish()


	# We read and send synchronously, so we can clean up safely after finish
	shutil.rmtree(tmp_dir, True)
		
        
        
        
        
        
        