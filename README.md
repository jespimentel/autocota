# autocota, by pimentel (2023)
## Gerador de cotas (em docx e PDF) para facilitar o trabalho com o eSAJ.

O programa, escrito em Python, facilita o trabalho na Promotoria de Justiça, gerando cotas nos formatos docx e pdf a partir de informações contidas em arquivo texto.

### Como utilizar
1. Copie para a área de transferência o conjunto das intimações que serão recebidas. Pode ser um simples "Control-A/Control-C" no navegador. O programa extrai os números dos processos.
2. Use a opção 1 para gerar o arquivo "cotas.txt" com as informações da área de transferência.
3. Edite o arquivo "cotas.txt" com o conteúdo das manifestações. 
4. Rode a aplicação mais uma vez. Certifique-se de que possui no mesmo diretório um documento base ("documento_base.docx"), com campos XXXX (onde será inserido o número do processo) e YYYY (onde serão inseridos os parágrafos). A área do texto a ser preenchida já deve estar formatada com o estilo desejado. 
5. Use a opção 2 para gerar os documentos docx e pdf automaticamente.

### Dicas
1. Confira regularmente as atualizações do programa em <https://github.com/jespimentel/autocota>.
2. Veja o programa funcionando: <https://youtu.be/QOd2AeEOhgU> (em menos de 2 minutos).

### Por baixo do capô

Principais bibliotecas utilizadas:

os: Fornece funcionalidades para interagir com o sistema operacional, como criação de diretórios, manipulação de caminhos de arquivos, etc.

re: Oferece suporte a expressões regulares para busca e manipulação de padrões de texto.

pyperclip: Permite acessar o conteúdo da área de transferência do sistema.

win32com.client: Fornece acesso às funcionalidades do Microsoft Office, como o Word neste caso.

datetime: Fornece classes para manipulação de datas e horas.

docx: Biblioteca para criar, ler e manipular arquivos do Word (.docx).
