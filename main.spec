
import gooey
gooey_root = os.path.dirname(gooey.__file__)
gooey_languages = Tree(os.path.join(gooey_root, 'languages'), prefix = 'gooey/languages')
gooey_images = Tree(os.path.join(gooey_root, 'images'), prefix = 'gooey/images')
a = Analysis([r'D:\Users\lj\Desktop\water\gui\main.py'],        # 项目文件名称
             pathex=[r'C:\Users\lj\AppData\Local\Programs\Python\Python38\Scripts'],  # python安装路径
             datas=[('chromedriver.exe','.')],
             hiddenimports=['matplotlib',
'django.contrib.admin.apps',
'django.contrib.auth.apps',
'django.contrib.gis.utils',
'django.contrib.gis.admin',
'django.contrib.contenttypes.apps',
'django.contrib.messages.apps',
'django.contrib.staticfiles.apps',
'django.contrib.sessions.models',
'django.contrib.sessions.apps',
'django.contrib.messages.middleware',
'django.contrib.auth.middleware',
'django.contrib.sessions.middleware',
'django.contrib.sessions.serializers',],
             hookspath=None,
             runtime_hooks=None,
             )
pyz = PYZ(a.pure)

options = [('u', None, 'OPTION'), ('u', None, 'OPTION'), ('u', None, 'OPTION')]

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          options,
          gooey_languages, # Add them in to collected files
          gooey_images, # Same here.
          name='mian.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False,
          windowed=True,
          icon=os.path.join(gooey_root, 'images', 'program_icon.ico'))