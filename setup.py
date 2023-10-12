import sys
import os
from cx_Freeze import setup, Executable

configuracao = Executable(
    script='main.py'    
)

setup(
    name='Busca Cartórios Brasil',
    version='2.0',
    description='Este programa faz uma busca de cartórios situados na cidade de determinado Estado informados pelo usuário.',
    author='Gabriel Alves Silva',
    options={'build_exee':{
        'include_msvcr': True,
    }},
    executables=[configuracao]
)