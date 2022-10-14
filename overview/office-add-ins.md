# Visão geral da plataforma de suplementos do Office

[Referência Externa](https://learn.microsoft.com/pt-br/office/dev/add-ins/overview/office-add-ins)

Use a plataforma de Suplementos do Office para criar soluções que estendem os aplicativos do Office e interagem com o conteúdo dos documentos do Office. Com os suplementos do Office, você pode usar tecnologias da Web conhecidas, como HTML, CSS e JavaScript para estender e interagir com o Outlook, Excel, Word, PowerPoint, OneNote e Project. Sua solução pode ser executada no Office em várias plataformas, insluindo no Windows, Mac, iPad e em um navegador.

![O aplicativo do Office mais um site inserido (suplemento) tornam infinitas as possibilidades de extensibilidade.](../assets/images/addins-overview.png)

Os suplementos do Office podem fazer quase tudo que uma página da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:

- **Adicione novas funcionalidades aos clientes do Office** - Traga dados externos para o Office, automatize documentos do Office, exponha funcionalidades da Microsoft e de outros clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar a dados que inpulsionam a produtividade.

- **Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** - Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint.

# Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO?

Os suplementos de COM ou VSTO são soluções de integração anteriores do Office que são executadas apenas no Office no Windows. Ao contrário de suplementos de COM, os suplementos do Office não envolvem código executada no dispositivo do usuário ou no cliente do Office. Para um suplemento do Office, o aplicativo do cliente (por exemplo, o Excel), lê o manifesto do suplemento e conecta os comandos do meno e os botões da faixa de opções personalizada do suplemento à interface de usuário. Quando necessário, ele carrega o código de HTML e o JavaScript, que são executados no contexto de um navegador em uma área restrita.

![Os motivos para usar os Suplementos do Office: multiplataforma, implantação centralizada, acesso fácil por meio do AppSource e baseado em tecnologias Web padrão.](../assets/images/why.png)

Os suplementos do Office oferecem as seguintes vantagens em relação aos suplementos criados usando VBA, COM ou VSTO.

- Suporte à plataforma cruzada. Os suplementos do Office podem ser executados no Office na Web, Windows, Mac e iPad.
- Implantação e distribuição centralizadas. Os administradores podem implantar suplementos do Office centralmente em uma organização.
- Acesso fácil através da AppSource. Você pode disponibilizar sua solução para um público amplo ao enviá-la para o AppSource.
- Com base na tecnologia da Internet padrão. Você pode usar qualquer biblioteca que gosta para criar suplementos do Office.