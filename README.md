# csvintoexcel

## 1. Introdução ##

Este repositório contém o código fonte do componente **csvintoexcel**. Este componente distribuído como um arquivo (.jar) pode ser executado em linha de comando (tanto Windows quanto Linux). O **csvintoexcel** junta o conteúdo de um arquivo de dados dentro de uma planilha Excel (.xls) ou (.xlsx) a partir de uma posição de linha-inicial e coluna-inicial (x,y).


### 1.1. Help Online linha

* O componente **csvintoexcel** funciona com argumentos de linha de comando (tanto Windows quanto Linux). Com o argumento '-h' mostra o help da a aplicação.

```bat
C:\My Git\workspace-github\csvintoexcel\dist>java -jar csvintoexcel.jar -h
csvintoexcel [v01.00.20190202] Tool to join excel template with csv content file.
  [-h|--help]
        Print help message

  [-e <input-excel-file>]
        Nome do arquivo que contém Pasta de trabalho EXCEL (.xls ou .xlsx) usada na juncao. Ex: Template.xlsx

  [-s <input-excel-sheet-number>]
        Numero sequencial da PLANILHA dentro da Pasta de trabalho usada na juncao. Ex: 1 (primeira aba. Default é primeira planilha)

  [-r <input-excel-row-number>]
        Numero da LINHA inicial da planilha usada na juncao. Ex: 2 (segunda linha. Default é segunda linha)

  [-c <input-excel-column-number>]
        Numero da COLUNA inicial da planilha. Ex: 1 (primeira coluna. Default é primeira linha)

  [-f <input-csv-file>]
        Nome do arquivo (.csv) que contém o conteúdo usado na juncao. Ex: dados.csv (separador ;)

  [-i <input-csv-file-ignore-header>]
        Numero de LINHAS DE CABECALHO dos dados ignorados na juncao. Ex: 1 (Default é 1)

Execution aborted!
```




### 2. Documentação ###

### 2.1. Diagrama de Caso de Uso (Use Case Diagram) ###

#### Diagrama de Caso de Uso

![UseCaseDiagram](doc/UseCaseDiagram%20-%20Context%20-csvintoexcel.png) 


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
