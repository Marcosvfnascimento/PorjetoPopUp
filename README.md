# PorjetoPopUp - Marcos Vinícius

# Lembrete de Contratos - ARPE

Este projeto em Python exibe alertas automáticos de contratos que estão próximos ao vencimento. Ele utiliza um arquivo Excel como base de dados e apresenta notificações via interface gráfica (Tkinter), permitindo marcar o andamento ou resolução do contrato diretamente na tela.

---

## Funcionalidades

- Verifica periodicamente contratos com menos de 90 dias para vencimento.
- Mostra alertas pop-up com detalhes do contrato.
- Permite atualizar o status diretamente pela interface.
- Atualiza automaticamente o arquivo Excel com os novos status.

---

## Exemplo de execução

Certifique-se de que os arquivos contratos.xlsx e ARPE.jpg estejam na mesma pasta do script lembrete_contratos_pop_up.py.

### No terminal execute o seguinte comando
python lembrete_contratos_pop_up.py


---

## Requisitos

Certifique-se de ter o Python 3.7+ instalado, além das seguintes bibliotecas:

```bash
pip install pandas schedule pillow

Estrutura esperada Excel (contratos.xlsx)
O arquivo deve conter, no mínimo, as seguintes colunas:

- Empresa

- ontrato

- Vigencia_Termino (data no formato reconhecido pelo Excel)

- Status (valores esperados: em branco, Em andamento, Resolvido)



