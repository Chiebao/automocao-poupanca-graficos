# App de Poupança com Gráficos

Aplicativo desktop em Python que registra quanto foi economizado a cada dia, guarda tudo numa planilha Excel e mostra a evolução em gráficos, sem precisar abrir a planilha.

---

## Funcionalidades

- Registro do valor economizado no dia, aceitando vírgula ou ponto
- Salvamento automático em `valores_diarios.xlsx` (a planilha é criada na primeira execução)
- Total acumulado atualizado na hora
- Gráfico de barras com a economia por mês e gráfico de pizza com a economia por semana, atualizados a cada novo valor
- Validação de entrada, com aviso quando o valor não é numérico

## Como executar

```bash
pip install openpyxl matplotlib
python main.py
```

## Tecnologias

`Python` `Tkinter` `openpyxl` `Matplotlib`

## Estrutura

```
/
├── main.py                # Interface, gravação no Excel e gráficos
├── valores_diarios.xlsx   # Planilha de dados
└── README.md
```

---

**Luana Godoy** · Estudante de Análise e Desenvolvimento de Sistemas
[LinkedIn](https://www.linkedin.com/in/luana-chiebao) · [GitHub](https://github.com/Chiebao)
