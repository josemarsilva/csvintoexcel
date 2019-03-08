REM ---------------------------------------------------------------------------
REM Filename: querytocsv-stdmassivafrontend
REM Author   : josemarsilva@inmetrics.com.br
REM LastRev  : 2019-02-14
REM ---------------------------------------------------------------------------

ECHO Off
ECHO.
ECHO "Setting up ..."
ECHO.
REM SET SED_HOME="C:\Inmetrics\gnu-win32-sed"
REM SET SED_HOME="E:\Inmetrics\gnu-win32-sed"
REM SET SED_HOME="C:\Program Files (x86)\GnuWin32\bin"
REM SET SED_HOME="E:\tools-inmetrics\gnu-win32-sed"
SET SED_HOME="E:\tools-inmetrics\gnu-win32-sed"

SET DT_BATCH=28\/02\/2019
SET NU_LOAD_FILE_ID=12864870
SET DT2_BATCH=27\/02\/2019
SET NU2_LOAD_FILE_ID=12854586
SET DT3_BATCH=%DT2_BATCH%
SET NU3_LOAD_FILE_ID=%NU2_LOAD_FILE_ID%

REM MaxRows 200k - NFiles 1 - EachFileMax 200k
SET ROWNUM_LIMIT=200000
SET PART_SIZE=100000
SET PART_NFILES=1

SET USERNAME=starstausr
SET PASSWORD=starstausr
SET HOSTNAME=10.64.113.16
SET PORT=1521
REM SET SID=BDTHSTAR
SET SID=BDTHSTAR1

SET YMD_DATE_SUFIX=%date:~6,4%%date:~3,2%%date:~0,2%

ECHO.
ECHO "Building temporary scripts from template ..."
ECHO.

%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.sql" > "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.tmp"
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "02 - STD-FRONTEND - CategoriaDeAntecipacao.sql"                 > "02 - STD-FRONTEND - CategoriaDeAntecipacao.tmp"                
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "03 - STD-FRONTEND - DebitosSemRetorno.sql"                      > "03 - STD-FRONTEND - DebitosSemRetorno.tmp"                     
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "04 - STD-FRONTEND - EfetivacaoDeAjuste.sql"                     > "04 - STD-FRONTEND - EfetivacaoDeAjuste.tmp"                    
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "05 - STD-FRONTEND - GoldListDeMarcas.sql"                       > "05 - STD-FRONTEND - GoldListDeMarcas.tmp"                      
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "06 - STD-FRONTEND - IncentivoDeAluguel.sql"                     > "06 - STD-FRONTEND - IncentivoDeAluguel.tmp"                    
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "07 - STD-FRONTEND - ReprocessamentoDeVendas.sql"                > "07 - STD-FRONTEND - ReprocessamentoDeVendas.tmp"               
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "08 - STD-FRONTEND - ReversaoDeCancelamento.sql"                 > "08 - STD-FRONTEND - ReversaoDeCancelamento.tmp"                
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "09 - STD-FRONTEND - CancelamentoDeVenda.sql"                    > "09 - STD-FRONTEND - CancelamentoDeVenda.tmp"                
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "10 - STD-FRONTEND - MdrMinimo.sql"                              > "10 - STD-FRONTEND - MdrMinimo.tmp"                
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "11 - STD-FRONTEND - Hierarquia.sql"                             > "11 - STD-FRONTEND - Hierarquia.tmp"                
%SED_HOME%\sed "s/${nu_load_file_id}/%NU_LOAD_FILE_ID%/g;s/${dt_batch}/%DT_BATCH%/g;s/${nu2_load_file_id}/%NU2_LOAD_FILE_ID%/g;s/${dt2_batch}/%DT2_BATCH%/g;s/${rownum_limit}/%ROWNUM_LIMIT%/g;s/${nu3_load_file_id}/%NU3_LOAD_FILE_ID%/g;s/${dt3_batch}/%DT3_BATCH%/g" "12 - STD-FRONTEND - EnvioDeDebito.sql"                          > "12 - STD-FRONTEND - EnvioDeDebito.tmp"                


ECHO.
ECHO "Query database to (.csv) file ..."
ECHO.

java -jar ../querytocsv.jar -f "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.tmp" -r "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.csv" -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "02 - STD-FRONTEND - CategoriaDeAntecipacao.tmp"                 -r "02 - STD-FRONTEND - CategoriaDeAntecipacao.csv"                 -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "03 - STD-FRONTEND - DebitosSemRetorno.tmp"                      -r "03 - STD-FRONTEND - DebitosSemRetorno.csv"                      -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "04 - STD-FRONTEND - EfetivacaoDeAjuste.tmp"                     -r "04 - STD-FRONTEND - EfetivacaoDeAjuste.csv"                     -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "05 - STD-FRONTEND - GoldListDeMarcas.tmp"                       -r "05 - STD-FRONTEND - GoldListDeMarcas.csv"                       -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "06 - STD-FRONTEND - IncentivoDeAluguel.tmp"                     -r "06 - STD-FRONTEND - IncentivoDeAluguel.csv"                     -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "07 - STD-FRONTEND - ReprocessamentoDeVendas.tmp"                -r "07 - STD-FRONTEND - ReprocessamentoDeVendas.csv"                -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "08 - STD-FRONTEND - ReversaoDeCancelamento.tmp"                 -r "08 - STD-FRONTEND - ReversaoDeCancelamento.csv"                 -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "09 - STD-FRONTEND - CancelamentoDeVenda.tmp"                    -r "09 - STD-FRONTEND - CancelamentoDeVenda.csv"                    -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "10 - STD-FRONTEND - MdrMinimo.tmp"                              -r "10 - STD-FRONTEND - MdrMinimo.csv"                              -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "11 - STD-FRONTEND - Hierarquia.tmp"                             -r "11 - STD-FRONTEND - Hierarquia.csv"                             -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%
java -jar ../querytocsv.jar -f "12 - STD-FRONTEND - EnvioDeDebito.tmp"                          -r "12 - STD-FRONTEND - EnvioDeDebito.csv"                          -t oracle -o jdbc:oracle:thin:%USERNAME%/%PASSWORD%@%HOSTNAME%:%PORT%:%SID%


ECHO.
ECHO "Counting rows (.csv) ..."
ECHO.
ECHO NomeArquivoCsv;QtdeLinhas > querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.csv" | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "02 - STD-FRONTEND - CategoriaDeAntecipacao.csv"                 | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 02 - STD-FRONTEND - CategoriaDeAntecipacao.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "03 - STD-FRONTEND - DebitosSemRetorno.csv"                      | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 03 - STD-FRONTEND - DebitosSemRetorno.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "04 - STD-FRONTEND - EfetivacaoDeAjuste.csv"                     | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 04 - STD-FRONTEND - EfetivacaoDeAjuste.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "05 - STD-FRONTEND - GoldListDeMarcas.csv"                     | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 05 - STD-FRONTEND - GoldListDeMarcas.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "06 - STD-FRONTEND - IncentivoDeAluguel.csv"                     | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 06 - STD-FRONTEND - IncentivoDeAluguel.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "07 - STD-FRONTEND - ReprocessamentoDeVendas.csv"                | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 07 - STD-FRONTEND - ReprocessamentoDeVendas.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "08 - STD-FRONTEND - ReversaoDeCancelamento.csv"                 | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 08 - STD-FRONTEND - ReversaoDeCancelamento.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "09 - STD-FRONTEND - CancelamentoDeVenda.csv"                    | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 09 - STD-FRONTEND - CancelamentoDeVenda.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "10 - STD-FRONTEND - MdrMinimo.csv"                              | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 10 - STD-FRONTEND - MdrMinimo.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "?.csv"                                                          | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 11 - STD-FRONTEND - Hierarquia.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "11 - STD-FRONTEND - Hierarquia.csv"                             | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 11 - STD-FRONTEND - Hierarquia.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv
findstr /R /N "^" "12 - STD-FRONTEND - EnvioDeDebito.csv"                          | find /C ":" > count.tmp
set /p COUNT=<count.tmp
ECHO 12 - STD-FRONTEND - EnvioDeDebito.sql;%COUNT% >> querytocsv-stdmassivafrontend.csv


ECHO.
ECHO Particiona arquivos (.csv) ...
ECHO.

SET part_size=%PART_SIZE%
SET /A i=1
:loop_part

  SET /A part_ini=(%i% - 1) * %part_size% + 2
  SET /A part_end=%i% * %part_size% + 1
  
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.csv" > "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida - %i% - %YMD_DATE_SUFIX%.csv"
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "02 - STD-FRONTEND - CategoriaDeAntecipacao.csv"                 > "02 - STD-FRONTEND - CategoriaDeAntecipacao - %i% - %YMD_DATE_SUFIX%.csv"                
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "03 - STD-FRONTEND - DebitosSemRetorno.csv"                      > "03 - STD-FRONTEND - DebitosSemRetorno - %i% - %YMD_DATE_SUFIX%.csv"                     
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "04 - STD-FRONTEND - EfetivacaoDeAjuste.csv"                     > "04 - STD-FRONTEND - EfetivacaoDeAjuste - %i% - %YMD_DATE_SUFIX%.csv"                     
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "05 - STD-FRONTEND - GoldListDeMarcas.csv"                       > "05 - STD-FRONTEND - GoldListDeMarcas - %i% - %YMD_DATE_SUFIX%.csv"                      
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "06 - STD-FRONTEND - IncentivoDeAluguel.csv"                     > "06 - STD-FRONTEND - IncentivoDeAluguel - %i% - %YMD_DATE_SUFIX%.csv"                    
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "07 - STD-FRONTEND - ReprocessamentoDeVendas.csv"                > "07 - STD-FRONTEND - ReprocessamentoDeVendas - %i% - %YMD_DATE_SUFIX%.csv"               
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "08 - STD-FRONTEND - ReversaoDeCancelamento.csv"                 > "08 - STD-FRONTEND - ReversaoDeCancelamento - %i% - %YMD_DATE_SUFIX%.csv"                
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "09 - STD-FRONTEND - CancelamentoDeVenda.csv"                    > "09 - STD-FRONTEND - CancelamentoDeVenda - %i% - %YMD_DATE_SUFIX%.csv"                   
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "10 - STD-FRONTEND - MdrMinimo.csv"                              > "10 - STD-FRONTEND - MdrMinimo - %i% - %YMD_DATE_SUFIX%.csv"                             
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "11 - STD-FRONTEND - Hierarquia.csv"                             > "11 - STD-FRONTEND - Hierarquia - %i% - %YMD_DATE_SUFIX%.csv"                            
  %SED_HOME%\sed -n "1p;%part_ini%,%part_end%p" "12 - STD-FRONTEND - EnvioDeDebito.csv"                          > "12 - STD-FRONTEND - EnvioDeDebito - %i% - %YMD_DATE_SUFIX%.csv"                         

  SET /A i=%i% + 1
  IF %i% LEQ %PART_NFILES% GOTO :loop_part

ECHO.
ECHO CSV INSERT INTO Excel..."
ECHO.

SET /A i=1
:loop_csvintoexcel

  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida.xls" -f "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida - %i% - %YMD_DATE_SUFIX%.csv" -o "01 - STD-FRONTEND - BaixaSistemicaIndependenteStatusDivida - %i% - %YMD_DATE_SUFIX%.xls" -d s-s-s-s-s-s-f-s
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/02 - STD-FRONTEND - CategoriaDeAntecipacao.xlsx"                -f "02 - STD-FRONTEND - CategoriaDeAntecipacao - %i% - %YMD_DATE_SUFIX%.csv"                 -o "02 - STD-FRONTEND - CategoriaDeAntecipacao - %i% - %YMD_DATE_SUFIX%.xlsx"                -d s
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/03 - STD-FRONTEND - DebitosSemRetorno.xlsx"                     -f "03 - STD-FRONTEND - DebitosSemRetorno - %i% - %YMD_DATE_SUFIX%.csv"                      -o "03 - STD-FRONTEND - DebitosSemRetorno - %i% - %YMD_DATE_SUFIX%.xlsx"                     -d s-----f---- -u UTF-8
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/05 - STD-FRONTEND - GoldListDeMarcas.xlsx"                      -f "05 - STD-FRONTEND - GoldListDeMarcas - %i% - %YMD_DATE_SUFIX%.csv"                       -o "05 - STD-FRONTEND - GoldListDeMarcas - %i% - %YMD_DATE_SUFIX%.xlsx"                      -r 2 -c 0 -d ------s---f-f-----
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/06 - STD-FRONTEND - IncentivoDeAluguel.xlsx"                    -f "06 - STD-FRONTEND - IncentivoDeAluguel - %i% - %YMD_DATE_SUFIX%.csv"                     -o "06 - STD-FRONTEND - IncentivoDeAluguel - %i% - %YMD_DATE_SUFIX%.xlsx"                    -r 8 -c 0
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/07 - STD-FRONTEND - ReprocessamentoDeVendas.xlsx"               -f "07 - STD-FRONTEND - ReprocessamentoDeVendas - %i% - %YMD_DATE_SUFIX%.csv"                -o "07 - STD-FRONTEND - ReprocessamentoDeVendas - %i% - %YMD_DATE_SUFIX%.xlsx"               -d s-s---f--------
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/08 - STD-FRONTEND - ReversaoDeCancelamento.xlsx"                -f "08 - STD-FRONTEND - ReversaoDeCancelamento - %i% - %YMD_DATE_SUFIX%.csv"                 -o "08 - STD-FRONTEND - ReversaoDeCancelamento - %i% - %YMD_DATE_SUFIX%.xlsx"                -d s--s---f-f
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/09 - STD-FRONTEND - CancelamentoDeVenda.xlsx"                   -f "09 - STD-FRONTEND - CancelamentoDeVenda - %i% - %YMD_DATE_SUFIX%.csv"                    -o "09 - STD-FRONTEND - CancelamentoDeVenda - %i% - %YMD_DATE_SUFIX%.xlsx"                   -r 11 -c 0 -d s-----f-f-s-
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/10 - STD-FRONTEND - MdrMinimo.xlsx"                             -f "10 - STD-FRONTEND - MdrMinimo - %i% - %YMD_DATE_SUFIX%.csv"                              -o "10 - STD-FRONTEND - MdrMinimo - %i% - %YMD_DATE_SUFIX%.xlsx"                             -d s-s
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/11 - STD-FRONTEND - Hierarquia.xlsx"                            -f "11 - STD-FRONTEND - Hierarquia - %i% - %YMD_DATE_SUFIX%.csv"                             -o "11 - STD-FRONTEND - Hierarquia - %i% - %YMD_DATE_SUFIX%.xlsx"                            -d s-s-s-s----------------
  java -jar ..\csvintoexcel.jar -e "./csvintoexcel-templates/12 - STD-FRONTEND - EnvioDeDebito.xlsx"                         -f "12 - STD-FRONTEND - EnvioDeDebito - %i% - %YMD_DATE_SUFIX%.csv"                          -o "12 - STD-FRONTEND - EnvioDeDebito - %i% - %YMD_DATE_SUFIX%.xlsx"                         -d s

  SET /A i=%i% + 1
  IF %i% LEQ %PART_NFILES% GOTO :loop_csvintoexcel


ECHO.
ECHO Fim!
ECHO.
ECHO On
