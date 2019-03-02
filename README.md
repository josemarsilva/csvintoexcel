# csvintoexcel

## 1. Introdução ##

Este repositório contém o código fonte do componente **csvintoexcel**. Este componente distribuído como um arquivo (.jar) pode ser executado em linha de comando (tanto Windows quanto Linux). O **csvintoexcel** junta o conteúdo de um arquivo de dados dentro de uma planilha Excel (.xls) ou (.xlsx) a partir de uma posição de linha-inicial e coluna-inicial (x,y).


### 1.1. Help Online linha

* O componente **csvintoexcel** funciona com argumentos de linha de comando (tanto Windows quanto Linux). Com o argumento '-h' mostra o help da a aplicação.

```bat
C:\My Git\workspace-github\csvintoexcel\dist>java -jar csvintoexcel.jar
Missing required options: e, f, o
usage: csvintoexcel [<args-options-list>] - v.2019.03.02
 -c,--input-excel-column-number <arg>      Numero da COLUNA inicial da
                                           planilha. Ex: 0-n (0=primeira
                                           coluna). Default 0
 -d,--input-excel-data-type-list <arg>     Lista dos tipos de dados
                                           (data-type) das celulas
                                           separados por '-'. Onde: 's':
                                           String; 'd': Numerico; 'f':
                                           decimais.
 -e,--input-excel-file <arg>               Nome do arquivo que contem
                                           Pasta de trabalho EXCEL (.xls
                                           ou .xlsx) usada na juncao. Ex:
                                           template.xlsx
 -f,--input-csv-file <arg>                 Nome do arquivo (.csv) que
                                           contem o conteudo usado na
                                           juncao. Ex: dados.csv
                                           (separador ;)
 -g,--input-csv-file-ignore-header <arg>   Numero de LINHAS DE CABECALHO
                                           dos dados ignorados na juncao.
                                           Ex: 1 (Default e 1)
 -h,--help                                 shows usage help message. See
                                           more
                                           https://github.com/josemarsilva
                                           /csvintoexcel
 -o,--output-excel-file <arg>              Nome do arquivo que da Pasta de
                                           trabalho EXCEL (.xls ou .xlsx)
                                           conter a juncao. Ex:
                                           template-com-dados.xlsx
 -r,--input-excel-row-number <arg>         Numero da LINHA inicial da
                                           planilha usada na juncao. Ex:
                                           0-n (0=primeira). Default 2
 -s,--input-excel-sheet-number <arg>       Numero sequencial da PLANILHA
                                           dentro da Pasta de trabalho
                                           usada na juncao. Ex: 0-n
                                           (0=primeira). Default 0
 -u,--output-excel-utf-encoding <arg>      UTF Encoding da planilha. Ex:
                                           UTF-8 (Default nenhum econding
                                           sera usado)
```




### 2. Documentação ###

### 2.1. Diagrama de Caso de Uso (Use Case Diagram) ###

#### Diagrama de Caso de Uso

![UseCaseDiagram](doc/UseCaseDiagram%20-%20Context%20-%20CsvIntoExcel.png)


### 2.5. Requisitos ###

* n/a


## 3. Projeto ##

### 3.1. Pré-requisitos ###

* Linguagem de programação: Java
* IDE: Eclipse (recomendado Oxigen 2)
* JDK: 1.8
* Maven Repository dependence: Apache CLI
* Maven Repository dependence: Apache POI

### 3.2. Guia para Desenvolvimento ###

* Obtenha o código fonte através de um "git clone". Utilize a branch "master" se a branch "develop" não estiver disponível.
* Faça suas alterações, commit e push na branch "develop".


### 3.3. Guia para Configuração ###

* n/a


### 3.4. Guia para Teste ###

* n/a


### 3.5. Guia para Implantação ###

* Obtenha o último pacote (.war) estável gerado disponível na sub-pasta './dist'.



### 3.6. Guia para Demonstração ###

* n/a


### 3.7. Guia para Execução ###

* Exemplo do uso do **csvintoexcel**  
    * Suponha um arquivo 'template.xlsx'.
    * Suponha um arquivo 'dados.csv'.

```bat
java -jar csvintoexcel.jar 

```


## Referências ##

* n/a
