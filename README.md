<h1>Calcular a diferença entre duas datas considerando dias e horas úteis.</h1>

### Créditos
Foi criado segundo explicação do canal Data Become no youtube.

✔️ Foi utilizado o recurso de parâmetros para definir o horário de expediente.

1) Dentro do power query, clique com o botão direito na área esquerda da tela (Onde fica as tabelas). E selecione <strong>Novo Parâmetro...</strong>
<img alt="Menu Parâmetros" title="Menu Parâmetros" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/menu-novo-parametro.png" />

2) Preencha o formulário, como na imagem abaixo, para criar o parâmetro de horario de inicio do expediente.
<img alt="Parâmetro Início Expediente" title="Parâmetro Início Expediente" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/parametro-inicio-expediente.png" />

3) Preencha o formulário, como na imagem abaixo, para criar o parâmetro de horario de fim do expediente.
<img alt="Parâmetro Início Expediente" title="Parâmetro Início Expediente" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/parametro-inicio-expediente.png" />

4) Foi criado o recurso de funções para descobrir a quantidade de dias úteis entre as datas. Para isso, clique com o botão direito em cima da tabela Chamados e selecione a opção <strong>Criar Função...</strong>
<img alt="Opção Criar Função" title="Opção Criar Função" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/opcao-criar-funcao.png" />

5) Dê um nome para a função desejada.
<img alt="Criar Função" title="Criar Função" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/criar-funcao.png" />

6) Selecione a função no canto esquerdo e clique em editor avançado.
<img alt="Criar Função" title="Criar Função" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/criar-funcao-002.png" />

7) Dentro do editor avançado insira a função já previamente criada.
```
/*
   A função tem como objetivo extrair a quantidade de horas úteis entre duas datas
   excluindo feriados informados como parâmetro e também um expediente ( hora de inicio e fim)

*/


(InicioExpediente, FimExpediente, Abertura, Fechamento, ListaFeriados) =>

let 

DiaDaAbertura   = Number.From(DateTime.Date(Abertura)),
DiaDoFechamento = Number.From(DateTime.Date(Fechamento)),

HorarioDaAbertura   = Number.From(DateTime.Time(Abertura)),
HorarioDoFechamento = Number.From(DateTime.Time(Fechamento)),

// Lista dos dias sem Sábados e Domingos
ListaDeDatas = List.Select({DiaDaAbertura..DiaDoFechamento}, each Number.Mod(_,7)>1),

// Lista dos dias sem Sábados, Domingos e Feriádos.
// Retorna apenas os números diferentes não existentes na tabela feriado, ou seja apenas não feriados.
ListaDiasUteis = List.Difference(ListaDeDatas,ListaFeriados),


SomaHorasUteis = 
        // Verifica se o dia da abertua é igual ao dia do fechamento
        if DiaDaAbertura = DiaDoFechamento then
                if DiaDaAbertura = List.First(ListaDiasUteis) then
                        // Verifica se o dia de abertura não é feriado. (DtAbertura = DtFechamento)
                        List.Median({InicioExpediente,FimExpediente,HorarioDoFechamento}) - List.Median({InicioExpediente,FimExpediente,HorarioDaAbertura})
                else 0
        else (
                if DiaDaAbertura = List.First(ListaDiasUteis) then
                        // Verifica se o dia da abertura é dia útil (DtAbertura <> DtFechamento)
                        FimExpediente - List.Median({InicioExpediente,FimExpediente,HorarioDaAbertura})
                else 0
        )
        +
        (       
                if DiaDoFechamento = List.Last(ListaDiasUteis) then 
                        // Verifica se o dia de fechamento é dia útil (DtAbertura <> DtFechamento)
                        List.Median({InicioExpediente,FimExpediente,HorarioDoFechamento}) - InicioExpediente
                else 0
        )
        +
        (
                //Soma to total de horas úteis excluindo (DiaAbertura, DiaFechamento, Feriados, Sábados e Domingos)
                List.Count(List.Difference(ListaDiasUteis,{DiaDaAbertura,DiaDoFechamento}))*(FimExpediente - InicioExpediente)
        )

in 

SomaHorasUteis
```

8) Feito isso, é necessário transformar o tipo dos dados na tabela feriados para o tipo <strong>Número Inteiro</strong>

9) Depois clique com o botão direito na coluna feriados e selecione <strong>Adicionar como Nova Consulta</strong> (para transformar os dados em uma lista)
<img alt="Transformar em lista" title="Transformar em lista" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/transformar-lista.png" />

10) Com a lista criada, selecione-a e depois clique novamente em editor avançado. Dentro do editor avançado crie um novo parametro para inserir a lista criado dentro de um buffer para ser mais performático durante o armazenamento, conforme imagem abaixo.
<img alt="Editor Avançado" title="Editor Avançado" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/lista-editor-avancado.png" />

11) Agora selecione a tabela Chamados e clique em editor avançado. Para inserir a transformação das datas em numeros inteiros
```
let
    Fonte = Excel.Workbook(File.Contents("C:\powerbi\powerbi-calculo-horas-uteis\Chamados.xlsx"), null, true),
    Tabela1_Table = Fonte{[Item="Tabela1",Kind="Table"]}[Data],

    /* Alteração a ser feita */
    Start = Number.From(InicioExpediente),
    End = Number.From(FimExpediente),

    #"Tipo Alterado" = Table.TransformColumnTypes(Tabela1_Table,{{"Data Abertura", type datetime}, {"Data Fechamento", type datetime}})
in
    #"Tipo Alterado"
```

12) Agora, selecionado a tabela chamados, clique em adicionar coluna com base em uma função personalizada. Preencha o formulário como a imagem abaixo.
<img alt="Adicionar Coluna" title="Adicionar Coluna" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/adicionado-coluna.png" />

13) Caso apareça o erro abaixo, selecione a tabela chamados e clique em editor avançado.
<img alt="Erro Adicionar Coluna" title="Erro Adicionar Coluna" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/erro-adicionar-coluna.png" />

14) Altere os dados conforme a imagem abaixo
```
let
    Fonte = Excel.Workbook(File.Contents("C:\powerbi\powerbi-calculo-horas-uteis\Chamados.xlsx"), null, true),
    Tabela1_Table = Fonte{[Item="Tabela1",Kind="Table"]}[Data],

    /* Alteração a ser feita */
    Start = Number.From(InicioExpediente),
    End = Number.From(FimExpediente),

    #"Tipo Alterado" = Table.TransformColumnTypes(Tabela1_Table,{{"Data Abertura", type datetime}, {"Data Fechamento", type datetime}}),
    #"Função Personalizada Invocada" = Table.AddColumn(#"Tipo Alterado", "DiffDiasUteis", each DiffDiasUteis(Start, End, [Data Abertura], [Data Fechamento], lista_feriados))
in
    #"Função Personalizada Invocada"
```

15) O erro deverá sumir e a coluna previamente com erro deverá mostrar agora os dados com os valores em dias úteis.
<img alt="Correção dos dados da coluna" title="Correção dos dados da coluna" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/correcao-dados-coluna.png" />

16) Agora selecione o novo campo criado e clique em adiciona coluna. Selecione o menu padrão e o submenu multiplicar.
<img alt="Adicionar nova coluna" title="Adicionar nova coluna" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/menu-padrao-multiplicar.png" />

17) Preencha os dados conforme imagem abaixo. (Como o dia tem 24 horas, inseria o valor a ser multiplicado)
<img alt="Adicionar coluna horas uteis" title="Adicionar coluna horas uteis" src="https://github.com/robsonlopesjr/powerbi-calculo-horas-uteis/blob/main/img/coluna-horas-uteis.png" />