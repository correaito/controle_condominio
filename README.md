<h1 align="center">CondoAccountPro</h1>
<h4 align="center">Sistema de gerenciamento de contas de condomínio para finalidade de estudo e treinamento em Python.</h4>

<p align="center">
<a href="#tecnologias"> Tecnologias</a> | <a href="#informacao-uso">Como Usar</a>
</p>

[View demo](#)

<p align="left"> <img src="https://komarev.com/ghpvc/?username=correaito&label=Project%20views&color=0e75b6&style=flat" alt="correaito" /> </p>

![imagem](https://img.shields.io/badge/-Python-orange) ![imagem](https://img.shields.io/badge/-Tkinter-black) ![imagem](https://img.shields.io/badge/-Pandas-brown) ![imagem](https://img.shields.io/badge/-Selenium-green)

<a id="tecnologias" class="anchor"></a>
### :rocket:  Tecnologias

------------
Esse projeto foi desenvolvido como um Projeto Pessoal, com as seguintes tecnologias:

- [Python](https://www.python.org/ "Heading link")
- [Tkinter](https://docs.python.org/3/library/tkinter.html/ "Heading link")
- [Pandas](https://pandas.pydata.org/ "Heading link")
- [Selenium](https://selenium-python.readthedocs.io/ "Heading link")

<a id="informacao-uso" class="anchor"></a>
### :information_source:  Como Usar
------------
Para executar este aplicativo, você precisará apenas clonar e executar em seu IDE. 

Da sua linha de comando:

    # Clone este repositório
    $ git clone https://github.com/correaito/controle_condominio.git
    
    # Vá para o repositório
    $ cd controle_condominio
    
Agora, para executar o script, dentro do PyCharm, abra o arquivo main.py, clique com o botão direito do mouse, e depois em "Run main.py", ou com <kbd>SHIFT</kbd> + <kbd>CTRL</kbd> + <kbd>F10</kbd>.

<a id="observacoes" class="anchor"></a>
### :loudspeaker:  Observações
------------

1. Nas linhas 182/186 do arquivo main.py, em send_keys é necessário alterar 'usuário/senha' para suas credenciais de acesso à area de cliente da Copel. Se você for usuário de outra cia (Cesp, Light, etc), a function 'pegar_valor_copel' não irá funcionar e deverá ser adaptado para o portal da sua região. 

2. No arquivo Calculos.xlsx as únicas informações que você poderá alterar serão as colunas com as descrições das despesas, os nomes dos moradores e os números de apto. Recomendo não adicionar novas linhas ou colunas, porém, caso seja feito, deverá adaptar para o pandas fazer a leitura dessas novas linhas/colunas.

3. Apos utilizar o programa, no arquivo Gera_Faturas.xlsm, bastará clicar no botão 'Gerar Faturas' para que a macro filtre somente as informações contendo valor maior que zero, gerando assim as faturas. 

4. Caso esse script não funcione, verifique, além das dependencias contidas em requirements.txt, se a versão do chromedriver desse repositório é compatível com a versão do Chrome que você está utilizando, versão do Windows e plataforma x86 ou x64
------------
Feito com ♥ por Alan Garmatter 👋 
