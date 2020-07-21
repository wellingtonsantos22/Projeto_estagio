Este programa tem como função a extração e organização de informações de currículos lattes,
assim como cálculos necessários ao acompanhamento e autoanálise de Programas de Pós-Graduação.
------------------------------------------------------------------------------
---------------------ÍNDICE---------------------

1) Instalação do Python no Windows
	1.1) Instalação das Bibliotecas
	1.2) Execução do Programa
2) Instalação do Python no Mac
	2.1) Instalação das Bibliotecas
	2.2) Execução do Programa
3) Instalação do Python e bibliotecas no Linux
	3.1) Execução do Programa
4) Como adicionar títulos à Base de Correções
5) Como alterar os valores dos estratos

*Os currículos baixados em xml deverão ser colocados na pasta LattesPlan.
-------------------------------------------------------------------------------


1) INSTALAÇÃO DO PYTHON NO WINDOWS

1. Entre na pasta "Windows" e execute o arquivo "python-3.7.5.exe"
	- abrirá uma janela de aviso..

2. Clique em "Executar"
	- abrirá a janela de instalação do Python..

3. Marque a checkbox "Add Python 3.7 to PATH" e depois clique em "Install Now"
	- abrirá uma janela pedindo autorização..

4. Clique em "Sim"
	- a instalação prosseguirá e será concluída..

5. Clique em "Disable path length limit"
	- abrirá uma janela pedindo autorização novamente..

6. Clique em "Sim"

7. Clique em "Close"

1.1) Instalação das Bibliotecas

1. Ainda na pasta "Windows", execute o arquivo "PackInstaller.bat"
	

1.2) Executando o programa

1. Na pasta LattesPlan, execute o arquivo "LattesPlan.Bat"

--------------------------------------------------------


2) INSTALAÇÃO DO PYTHON NO MAC

1. Entre na pasta "Mac" e execute o arquivo python.3.5.7.pkg
	- abrirá a janela de instalação do Python..

2. Clique em "Continue" por 3 vezes
	- abrirá uma janela de permissão..

3. Clique em "Agree"

4. Clique em "Continue"

5. Clique em "Install"
	- Se você usar senha, o programa pedirá neste etapa para confirmar sua identidade.

6. Clique em "close"

2.1) Instalação das Bibliotecas

1. Abra o terminal

2. Digite os comando abaixo e aperte "enter", um de cada vez: 
	cd Desktop/LattesPlan
	chmod +x Mac/PackInstaller.sh
	./PackInstaller.sh

*Se a pasta "LattesPlan" não estiver no Desktop, basta trocar "Desktop/" pelo caminho correto.

-Packs Instalados. Pronto para uso.

2.2) Executando o programa

1. Ainda no terminal, digite o comando abaixo e aperte "enter":
	python LattesPlan.py

*Caso você tenha fechado o terminal no fim do passo anterior, deverá especificar o caminho:
	"Desktop/LattesPlan/LattesPlan.py"
--------------------------------------------------------


3) INSTALAÇÃO DO PYTHON E BIBLIOTECAS NO LINUX

1. Abra o terminal e digite o comando abaixo e aperte "enter":
	sudo apt-get install python

2. Finalizando a instalação, digite o comando abaixo e aperte "enter":
	sudo apt install python-pip

3. Digite os comandos abaixo e aperte "enter", um de cada vez:
	cd LattesPlan
	chmod +x Linux/PackInstaller.sh
	./PackInstaller.sh

*Se a pasta "LattesPlan" não estiver na pasta principal, especifique o
 caminho correto (Caminho/) antes de "LattesPlan", no comando "cd LattesPlan".

-Packs Instalados. Pronto para uso.

3.2) Executando o programa

1. No terminal, digite o comando abaixo e aperte "enter":
	python LattesPlan.py

*Caso você tenha fechado o terminal no fim do passo anterior, deverá especificar o caminho:
	"LattesPlan/LattesPlan.py"
	
---------------------------------------------------------

4) COMO ADICIONAR TÍTULOS À BASE DE CORREÇÕES

1. Clique com o botão direito sobre o arquivo "BaseDeCorecoes.txt"

2. Coloque em "Abrir com" e selecione "Bloco de Notas"

3. A listaBase, presente no arquivo, é adicionada à lista dos resultados onde são salvos os títulos
   das publicações para a contagem, necessárias para posteriormente ser feita a divisão da pontuação
   da publicação entre os coautores do mesmo PPG.

4. O mecanismo funciona como no exemplo abaixo:
	Exemplo:
		listaBase = ['título 1',
			    'título um',
			    'título 2',
			    'title 2',
			    'title 2'
			    ]

- Para esse exemplo, 2 professores cadastraram o 'título 1' em seus respectivos currículos,
o primeiro cadastrou como 'título 1' e o segundo como 'título um', então foi adicionado à
lista um de cada, para quando for feita a contagem de quantos 'título 1' e 'título um' existe
na lista resultado, contar com mais esses, assim corrigindo a quantidade.

- O mesmo vale para os 3 professores que cadastraram o 'título 2' em seus currículos, dois cadastraram
como 'título 2' (em portugês) e o terceiro cadastrou 'title 2' (em inglês), então é adicionada uma vez
o 'título 2' (pois o mesmo já está 2 vezes escrito desta forma na lista dos resultados) e duas vezes o
'title 2' (pois o mesmo já está escrito 1 vez desta forma na lista dos resultados).

- Então a listaBase será adicionada à lista dos resultados, o 'título 1' e 'título um' serão contados 2
vezes e o 'título 2' e 'title 2' serão contados 3 vezes.

5. Para adicionar um título, basta, ao final do último escrito, adicionar uma vírgula e pular uma linha (Enter),
então apertar a tecla "Tab" até que que fique na mesma direção que os títulos acima e escrever o título entre
aspas. Faça isso para os dois títulos que estão escritos diferentes. Como no exemplo abaixo:
	Exemplo:
		listaBase = ['título 1',
			    'título um',
			    'título 2',
			    'title 2',
			    'title 2',
			    'título 3',
			    'title three'
			    ]
	
	Um autor cadastrou como 'título 3' e o outro como 'title three', então os títulos foram adicionados
	à listaBase.
---------------------------------------------------------

5) Como alterar os valores dos estratos

1. As variáveis onde estão atribuídos os valores dos estratos estão no mesmo arquivo descrito acima.
   Basta apagar o valor que está após o sinal de igual e escrever o novo.