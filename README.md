# VBA-Email
## Planilha automatizada de envio de e-mails

### SOBRE:

Esse projeto surgiu através da necessidade de acelerar uma das partes do processo de contato com empresas que possuiam relações . Por se tratar de uma atividade corrqiqueira, no qual todos os meses eram realizados contatos padronizados via e-mail se tratando de transaçãoes bancárias, confirmações de informações, entre outros. Foi percebido que a parte de envio dos e-mails era despendido um tempo maior que o necessário, assim, tendo a possibilidade de otimizar o tempo.

![image](https://user-images.githubusercontent.com/100227422/163737405-01b14396-b768-47d7-be19-76afa91e36da.png)


### OBJETIVO: Criar através de uma planilha um sistema automatizado de envio de E-mails.

***Obs***: 

-Algumas mudanças de estrutura foram feitas para que o projeto possa se estender a outros usos e não necessariamente só para a empresa mencionada;

- Os dados contidos na planilha e nas imagens são todos fictícios, utilizados apenas para efeitos demonstrativos.


### RESULTADO OBTIDO: Diminuição do tempo de confecção e envio de e-mails de 2-3 dias para 1 dia.

### ESTRUTURA


---- Index do Banco de Dados ----


| Index       | Descrição |
|-------------|-----------|
| Status      | Status de negociação com a empresa | 
| ID          | ID de cadastro da empresa no banco de dados | 
| CNPJ/CPF    | Cadastro Nacional de Empresa Jurídica da empresa | 
| SIGLA       | Sigla da empresa | 
| NOME        | Nome da empresa | 
| ENDEREÇO    | Endereço da Localização da Empresa | 
| COMPLEMENTO | Complemento do Endereço | 
| CEP         | Código de Endereçamento Postal da Empresa | 
| CIDADE      | Cidade Onde a Empresa Está Cadastrada | 
| CONTATO     | Contato do Responsável Pela Empresa  | 
| CARGO       | Cargo do Contato do Responsável Pela Empresa | 
| E-MAIL      | E-mail de Contato Com a Empresa | 
| TELEFONE    | Telefone de COntato da Empresa  | 
| ENTRADA     | Data de Cadastro da Empresa | 

---- Arquivos ----


| Arquivos    | Descrição |
|---------------|-----------|
| userform.bas  | Códigos em VBA do Cadastro das Empresas    |   
| sendEmail.bas | Macros de Envio de E-mail          | 


### IMAGENS 

![image](https://user-images.githubusercontent.com/100227422/163737461-575a7475-81df-4d5e-8e49-7cf70612cac2.png)
Home Page

![image](https://user-images.githubusercontent.com/100227422/163737476-ce9c0f07-857a-43f6-be75-34204194601d.png)
Tela de Cadastro

![image](https://user-images.githubusercontent.com/100227422/163737554-6c7029c8-39b2-4bbd-990e-291616e3b97e.png)
Banco de Dados






