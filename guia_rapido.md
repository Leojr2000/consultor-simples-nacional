# 🚀 Guia Rápido - Consultor Simples Nacional

## ⏱️ 5 Minutos para Começar

### 1️⃣ Instalar Python (se não tiver)
- Download em: https://www.python.org/downloads/
- **Importante**: Marque "Add Python to PATH"

### 2️⃣ Instalar dependências
Abra o terminal/prompt e execute:
```bash
pip install -r requirements.txt
```

### 3️⃣ Preparar os dados
1. Abra `dados_cnpj.xlsx`
2. Coloque seus CNPJs na coluna A (um por linha)
3. Salve o arquivo

**Formatos aceitos:**
- ✅ `11.222.333/0001-81` (com formatação)
- ✅ `11222333000181` (sem formatação)

### 4️⃣ Executar
```bash
python consultor_cnpj.py
```

### 5️⃣ Abrir resultado
Abra o arquivo gerado: `relatorio_simples_nacional.xlsx`

---

## 📊 Entender o Resultado

| Coluna | O que significa |
|--------|---|
| **CNPJ** | CNPJ consultado (formatado) |
| **Razão Social** | Nome oficial da empresa |
| **Simples Nacional** | ✅ SIM / ❌ NÃO / ⚠️ ERRO |
| **Situação** | ATIVO, INATIVO, etc |
| **Data Abertura** | Quando a empresa foi registrada |
| **Erro** | Motivo (se houver problema) |

---

## 🎨 Cores no Excel

| Cor | Significado |
|-----|---|
| 🟢 **Verde** | Optante pelo Simples Nacional |
| 🔴 **Vermelho** | NÃO é optante |
| 🟡 **Amarelo** | Erro (CNPJ inválido, não encontrado, etc) |

---

## ❓ Problemas Comuns

### "ModuleNotFoundError: No module named 'openpyxl'"
**Solução:**
```bash
pip install openpyxl requests pandas
```

### "ConnectionError" ou "Timeout"
- Verifique sua conexão de internet
- Tente novamente em alguns minutos
- A API pode estar temporariamente indisponível

### CNPJ não encontrado
- Verifique se o CNPJ está correto
- A API pode estar com dados desatualizados
- Aguarde algumas horas e tente novamente

### Excel não abre o arquivo gerado
- Tente usar outro leitor (LibreOffice Calc, Google Sheets)
- Instale/atualize o Office

---

## ⚡ Dicas Úteis

1. **Volume grande (1000+)?**
   - Execute em lotes menores
   - Use delay maior: `consultor = ConsultorCNPJ(delay=2)`

2. **Precisa só de um CNPJ?**
   - Coloque apenas um no Excel e execute
   - Ou use o exemplo: `python exemplo_uso.py`

3. **Quer automatizar?**
   - Rode via agendador de tarefas (Windows/Mac/Linux)
   - Crie um script que chama `consultor_cnpj.py`

4. **Compartilhar resultado?**
   - Arquivo Excel pode ser enviado diretamente
   - Ou exporte como PDF a partir do Excel

---

## 🔄 Fluxo de Uso

```
┌─────────────────────┐
│  Seu arquivo XLSX   │ (com CNPJs na coluna A)
└──────────────┬──────┘
               │
               ▼
┌─────────────────────────────┐
│  python consultor_cnpj.py   │ (executa a consulta)
└──────────────┬──────────────┘
               │
               ▼
┌──────────────────────────────────┐
│ relatorio_simples_nacional.xlsx  │ (resultado formatado)
└──────────────────────────────────┘
```

---

## 📞 Suporte

- **Dúvidas sobre CNPJs?** Consulte a Receita Federal
- **Problemas com código?** Abra uma issue no GitHub
- **API fora?** Tente a API alternativa no código

---

**Pronto! Você tem tudo que precisa para consultar CNPJs 🎉**

Para mais detalhes, consulte o `README.md`
