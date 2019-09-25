using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo {
    class Program {
        static void Main (string[] args) {
            #region criacao de documento
                Document exemploDoc = new Document ();
            #endregion

            #region Criação da secao no documento
                 //Adiciona uma seção com o nome secaoCapa ao documento
                 //Cada secao pode ser entendida como uma pagina do documento
                 Section secaoCapa = exemploDoc.AddSection ();
            #endregion

            #region Criar um paragrafo
                //Cria um paragrafo com o nome titulo e adiciona seção secaoCapa
                //Os paragrafos são necessários para inserçao de textos imagens tabelas, etc
                Paragraph titulo = secaoCapa.AddParagraph ();
            #endregion

            #region Adiciona texto ao paragrafo
                titulo.AppendText("Exemplo de titulo\n\n");
            #endregion

            #region Formatar paragrafo
                //Através da propriedade HorizontalAlignment, é possivel alinhar o parágrafo
                titulo.Format.HorizontalAlignment=HorizontalAlignment.Center;

                //Cria um estilo com o nome estilo01 e adiciona ao documento
                ParagraphStyle estilo01=new ParagraphStyle(exemploDoc);

                //Adiciona um nome ao estilo01
                estilo01.Name="Cor do titulo";

                //Definir a cor do titulo
                estilo01.CharacterFormat.TextColor=Color.DarkCyan;

                //Define que o texto será em negrito
                estilo01.CharacterFormat.Bold=true;

                //Adiciona o estilo01 ao documento exemploDoc
                exemploDoc.Styles.Add(estilo01);

                //Aplica o estilo01 ai paragrafo titulo
                titulo.ApplyStyle(estilo01.Name);
            #endregion

            #region Trabalhar com tabulação
                //Adiciona um paragrafo textoXapa a seção secaoCapa
                Paragraph textoCapa = secaoCapa.AddParagraph();

                //Adiciona um texto ao paragrafo com tabulação
                textoCapa.AppendText("\tEste é um exemplo de texto com tabulação\n");

                //Adiciona um novo parágrado a mesma seção (secaoCapa)
                Paragraph textoCapa2 = secaoCapa.AddParagraph();

                textoCapa2.AppendText("\tBasicamente, representa uma pagina do documento e os paragrafos dentro de uma mesma seção"+"Obviamente, aparecem na mesma página");
            #endregion
            
            #region Inserir imagens
                //Adiciona um paragrafo a seção capa
                Paragraph imagemCapa= secaoCapa.AddParagraph();

                //Adiciona um texto ao paragrafo imagemCapa
                imagemCapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao documento\n\n");

                //Centraliza horizontalmente o parágrafo imagemCapa
                imagemCapa.Format.HorizontalAlignment=HorizontalAlignment.Center;

                //Adiciona uma imagem com o nome imgaemExemplo ao parágrafo imagemCopa
                DocPicture imagemExemplo=imagemCapa.AppendPicture(Image.FromFile(@"saida/imagens/logo_csharp.png"));

                //Define uma largura e uma altura para a imagem
                imagemExemplo.Width=300;
                imagemExemplo.Height=300;
            #endregion

            #region Adicionar nova seção
                //Adiciona uma nova seção
                Section sessaoCorpo= exemploDoc.AddSection();

                //Adiciona paragrafo nessa sessão, para que consiga inserir texto
                Paragraph paragrafoCorpo1=secaoCapa.AddParagraph();

                paragrafoCorpo1.AppendText("\tEste é um exemplo de parágrafo criado em uma nova seção."+"Como foi criada em uma nova seção, perceba que este texto aparece em uma noba página.");
            #endregion

            #region adicionando tabela
                //Adiciona uma tabela a seção secaoCorpo
                Table tabela= sessaoCorpo.AddTable(true);

                //Cria cabeçalho da tabela
                String[] cabecalho={"Item","Descrição","QTD","Preço Unit.","Preço"};

                //Cria dados da tabela
                String[][] dados={
                    new String[]{"Cenoura","Vegetal que faz bem para os olhos","4","R$4,00", "R$4,00"},

                    new String[]{"Beterraba","Vegetal que faz bem para os ossos","3","R$5,00", "R$5,00"},

                    new String[]{"Batata","Vegetal que faz bem para os musculos","2","R$3,00", "R$3,00"},

                    new String[]{"Tomate","Tomate fruta vermelha","1","R$6,00", "R$6,00"},
                };

                //Celulas na tabela
                tabela.ResetCells(dados.Length+1,cabecalho.Length);

                //Adiciona uma linha na posição 0 no vetor de linhas
                //E define que esta linha é o cabeçalho
                TableRow Linha1= tabela.Rows[0];
                Linha1.IsHeader=true;

                // Define a altura das linhas
                Linha1.Height = 23;

                // Formataçção do cabeçalho
                Linha1.RowFormat.BackColor=Color.AliceBlue;

                //Percorre as colunas do cabeçalho
                for (int i = 0; i < cabecalho.Length; i++)
                {
                    //Alinhamento das celulas
                    Paragraph p =Linha1.Cells[i].AddParagraph();
                    Linha1.Cells[i].CellFormat.VerticalAlignment=VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment=HorizontalAlignment.Center;

                    //Formatação dos textos do cabeçalho
                    TextRange TR= p.AppendText(cabecalho[i]);
                    TR.CharacterFormat.FontName="Arial";
                    TR.CharacterFormat.FontSize=14;
                    TR.CharacterFormat.TextColor=Color.Pink;
                    TR.CharacterFormat.Italic=true;
                }
                
                //Adiciona as linhas do corpo da tabela
                for (int r = 0; r < dados.Length; r++)
                {
                    TableRow LinhaDados = tabela.Rows[r + 1];
                    //Altura da linha
                    LinhaDados.Height=20;
                    
                    //Este FOR vai percorrer as colunas do vetor
                    for (int c = 0; c < dados[r].Length; c++)
                    {
                        LinhaDados.Cells[r].CellFormat.VerticalAlignment=VerticalAlignment.Middle;

                        //Preencher os dados nas linhas
                        Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                        
                        TextRange TR2= p2.AppendText(dados[r][c]);

                        //Formata as celulas
                        p2.Format.HorizontalAlignment=HorizontalAlignment.Center;
                        TR2.CharacterFormat.FontSize=12;
                        TR2.CharacterFormat.FontName="Arial";
                        TR2.CharacterFormat.TextColor=Color.Black;
                        TR2.CharacterFormat.Bold=true;
                    }
                }
            #endregion

            #region Salvar arquivo
                //Salvar arquivo em .DOCX
                //Utiliza o metodo saveTofile para salvar o arquivo no formato desejado
                //Assim como no Word, caso já exista um arquivo com este nome, é substituido 
                exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
            #endregion
        }
    }
}