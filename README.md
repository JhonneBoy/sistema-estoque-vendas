# Sistema de Estoque e Vendas

Sistema ERP de Estoque e Vendas
Descrição

Este projeto é um ERP simples para gerenciamento de estoque, vendas e emissão de notas fiscais. Desenvolvido em Python, utilizando Tkinter, ttkbootstrap, Pandas e ReportLab, ele permite:

- Gerenciar produtos, vendas e vendedores.
- Controlar estoque com alertas para quantidades baixas.
- Visualizar gráficos de dashboard de vendas e estoque.
- Gerar PDFs de notas fiscais usando um modelo personalizável.

---

## 🚀 Funcionalidades
- Cadastro de produtos  
- Controle de estoque com quantidades atualizadas  
- Registro de vendas 
- Geração e gestão de NF 
- Backup automático  
- Operações com planilhas Excel  
- Interface simples em Python

---

## 🗂 Bibliotecas principais
Bibliotecas principais:
- pandas – manipulação de dados.
- ttkbootstrap – interface gráfica moderna.
- matplotlib – gráficos para dashboard.
- reportlab – geração de PDFs.
- openpyxl – leitura e escrita de arquivos Excel.

---

## 🗂 Como Executar
 - python app.py
 ou
 - python app_aprimorado.py
---

## 🗂 Login - Credenciais padrão
- login: admin
- senha: 1234
---

## 🗂 Utilize as abas para

- Utilize as abas para:
- Gerenciar Produtos
- Gerenciar Vendas
- Gerenciar Vendedores
- Visualizar o Dashboard

---

## 🗂 Utilize as abas para
```bash
Funcionalidades Principais
|
└─ Produtos
    |
    ├─ Adicionar, editar e excluir produtos.
    |   └─ Controle de estoque com alerta visual para produtos com quantidade baixa.
    |    
    └─Vendas
    |   ├─ Registrar vendas com vinculação de produtos e vendedores.
    |   └─ Atualização automática do estoque após cada venda.
    |
    ├─ Vendedores
    |   ├─ Cadastro de vendedores com informações de contato.
    |   └─ Preenchimento automático em vendas vinculadas.
    |
    └─Dashboard
        ├─ Gráficos de estoque atual e vendas totais.
        ├─ Visualização rápida de produtos com estoque baixo.
        ├─ Nota Fiscal
        ├─ Geração de PDF de nota fiscal usando o modelo nota-modelo.png.
        ├─ Número NF, Série, Data
        ├─ CNPJ emitente e destinatário
        ├─ CFOP, NCM
        └─ Quantidade, Valor Unitário, ICMS, IPI, Frete, Placa
```
---

## Build com PyInstaller

Caso queira gerar o executável do projeto:
pyinstaller --onefile app.py

Arquivos gerados aparecerão na pasta build/ conforme a estrutura acima.

⚠️ Observação: Arquivos maiores que 50 MB podem precisar de Git LFS ao subir para o GitHub.

---

## Observações

- O Excel (produtos.xlsx) é obrigatório para inicialização do sistema.

- Notas fiscais são salvas no diretório do projeto automaticamente após a geração.

- Modelo de nota fiscal (nota-modelo.png) pode ser atualizado para refletir o layout desejado.

## 🗂 Estrutura do Projeto
```bash
meuprojeto/
├─ app.py                  # Código principal do ERP
├─ app_aprimorado.py       # Versão aprimorada do app
├─ NotaFiscal_*.pdf        # PDFs gerados de notas fiscais
├─ nota-modelo.png         # Modelo de nota fiscal para geração de PDFs
├─ requirements.txt        # Dependências do projeto
├─ README.md               # Documentação do projeto
├─ app.spec                # Arquivo de configuração do PyInstaller
├─ build/                  # Build gerado pelo PyInstaller
│   └─ app/
│       ├─ EXE-00.toc
│       ├─ PKG-00.toc
│       ├─ PYZ-00.pyz
│       ├─ PYZ-00.toc
│       ├─ app.pkg
│       ├─ base_library.zip
│       ├─ warn-app.txt
│       └─ xref-app.html
├─ produtos.xlsx           # Arquivo Excel com produtos, vendas e vendedores

Dependências

Instale todas as bibliotecas necessárias usando:
├─ pip install -r requirements.txt
