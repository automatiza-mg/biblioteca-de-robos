# Subfluxo que visa categorizar a decisão da LTS (concedida ou indeferimento) e exaurir todas as causas de indeferimento, para correlacionar com os códigos de cada indeferimento no SISAP. Atentar para as variações de preenchimento da lista de causas, p. ex.: (X), ( X ), (X ), etc
Text.CropText.CropTextBetweenFlags Text: texto_bim FromFlag: $'''LICENÇA PARA TRATAMENTO DE SAÚDE INDEFERIDA POR:''' ToFlag: $'''RESTRIÇÃO AO PORTE DE ARMAS:''' IgnoreCase: False CroppedText=> indeferimento IsFlagFound=> IsFlagFound
IF NotContains(indeferimento, $'''x''', True) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''CONCEDIDA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )PERDA DE PRAZO LEGAL''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''PERDA DE PRAZO LEGAL''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )VÍCIO NO ATESTADO ENVIADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''VÍCIO NO ATESTADO ENVIADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )DOCUMENTAÇÃO INCOMPLETA''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''DOCUMENTAÇÃO INCOMPLETA''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )ATESTADO SEM CID''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO SEM CID''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''NÃO SE TRATAR DE LICENÇA PARA TRATAMENTO DE SAÚDE''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )SE TRATAR DE ATESTADO PARA TERCEIROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''SE TRATAR DE ATESTADO PARA TERCEIROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) OUTROS''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''OUTROS''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X) ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X ) ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X )ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X)ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X) ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''( X)ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X ) ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE IF Contains(indeferimento, $'''(X )ATESTADO PRÉ-DATADO''', False) THEN
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''ATESTADO PRÉ-DATADO''' Column: $'''q''' Row: linha_loop
ELSE
    Excel.WriteToExcel.WriteCell Instance: excel_pericia Value: $'''CONFERIR MOTIVO INDEFERIMENTO''' Column: $'''q''' Row: linha_loop
END
