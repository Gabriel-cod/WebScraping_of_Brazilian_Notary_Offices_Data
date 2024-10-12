# Scrapy_Dados_de_Cartorios - Freelancer
## Objetivo
#### Desenvolver um sistema de coleta de dados de todos os cartórios das cidade solicitadas pelo usuário e armazenar as informações coletadas em um arquivo EXCEL separando os cartórios de cada cidade por planilhas dentro do mesmo arquivo.
## Procedimentos
#### Inicialmente foi desenvolvido um sistema para coleta de informações para acesso às cidades de interesse. Seria feita uma solicitação ao usuário de qual Estado as cidades requeridas pertencem, quais cidades desse Estado seriam coletadas as informações e qual o nome do arquivo excel que esses dados seriam armazenados.
#### Após esse processo de perguntas ao usuário, foi realizado um sistema de programação, utilizando o framework Selenium, para acessar o site indicado pelo cliente e seguir os passos para termos contato às informações de cartórios de cada cidade indicada a princípio. Feito isso, seria possível a coleta de cada informação exigida de cada cartório listado, já armazenando-os na planilha da cidade em análise no arquivo EXCEL denominado pelo beneficiário no início da automação.
#### Ao final de todos esses procedimentos, seria informado ao usuário que todas as extrações foram realizadas, e feito o questionamento se o cliente gostaria de realizar outra extração.
#### Durante todos esses processos, foi necessário alguns tratamentos de erros para caso fosse feita uma má utilização do programa pelo usuário, como em casos de informar nomes errados de Estado ou cidades.
## Guia do Usuário
#### Para o funcionamento correto do programa é necessário que o usuário tenha o Google Chrome, o Selenium e o Openpyxl. Em contra partida, há uma excessão para essa condição, ao disponibilizar um arquivo executável para o cliente, certifiquei-me que o usuário necessitaria somente da instalação do Google Chrome para facilitar sua utilização. 
#### Em casos de não possuir o executável, o beneficiário deve clonar esse repositório. Após isso, o usuário deve utilizar o arquivo "main.py" para sua execução, iniciando assim os processos citados anteriormente.
