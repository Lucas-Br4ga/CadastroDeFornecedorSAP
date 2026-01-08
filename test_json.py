"""
Script de teste para verificar a estrutura do JSON gerado
"""
import json
from pathlib import Path

# Caminho do JSON limpo
json_path = Path(__file__).parent / "Arquivos" / "fornecedor_limpo.json"

print(f"Verificando arquivo: {json_path}")
print(f"Arquivo existe: {json_path.exists()}\n")

if json_path.exists():
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    print("=" * 60)
    print("ESTRUTURA DO JSON")
    print("=" * 60)
    
    for key, value in data.items():
        print(f"\nChave: '{key}'")
        print(f"  Tipo: {type(value)}")
        
        if isinstance(value, dict):
            print(f"  Subchaves: {list(value.keys())}")
            for subkey, subvalue in value.items():
                print(f"    - {subkey}: {subvalue} (tipo: {type(subvalue)})")
        else:
            print(f"  Valor: {value}")
            print(f"  ⚠️  PROBLEMA: Deveria ser dict, mas é {type(value)}")
    
    print("\n" + "=" * 60)
    print("JSON COMPLETO:")
    print("=" * 60)
    print(json.dumps(data, indent=2, ensure_ascii=False))
else:
    print("❌ Arquivo não encontrado!")
    print("\nCrie um JSON de teste com:")
    print("""
{
    "empresa": {
        "razao_social": "Teste",
        "nome_fantasia": "Teste"
    },
    "endereco": {},
    "contato": {},
    "bancario": {},
    "geral": {}
}
    """)