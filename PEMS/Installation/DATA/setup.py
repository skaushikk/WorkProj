import distutils.core

distutils.core.setup(
    name='OrcFxAPI',
    version='1.0',
    description='Python interface to the OrcaFlex API',
    long_description='Python interface to the OrcaFlex API',
    license='Commercial',
    author='Orcina',
    author_email='orcina@orcina.com',
    url='http://www.orcina.com',
    platforms='Windows',
    py_modules=('OrcFxAPI', 'OrcFxAPIConfig')
)
