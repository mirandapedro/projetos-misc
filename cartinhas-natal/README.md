# cartinhas-natal
 Projeto para gerar automaticamente as cartinhas da campanha de Natal da Casa do Caminho.

 O arquivo "medidas.ods" precisa ter os seguintes campos de cada criança:
 - "turma" - opcional, apenas para controle 
 - "criança" - o nome da criança
 - "idade" 
 - "calçado" - tamanho do calçado
 - "camisa" - tamanho da camisa
 - "calça" - tamanho da calça
 - "sexo"


O arquivo "docs\modelo.docx" já tem quase todos os detalhes prontos do modelo da cartinha. Programa apenas acrescenta a foto e dados de cada criança.

Na pasta "fotos" precisa estar a foto de cada criança, e o nome do arquivo dever ser "_Nome da Criança.jpg".

Em seguida, para cada criança, na pasta cartinhas, é gerado um arquivo docx, um pdf e um png (a partir do pdf). É necessário consertar a conversão para pdf - estava dando muitos erros quando utilizei.

Além disso, no final do script, é gerado um arquivo "etiquetas.docx", que contém a foto e o nome das crianças em sequência para facilitar a impressão de etiquetas para a identificação dos presentes. Por enquanto, arremate visual final do arquivo das etiquetas está sendo feito manualmente. Também é necessário adicionar o programa para adicionar coluna nas etiquetas com numeração única do presente/criança.