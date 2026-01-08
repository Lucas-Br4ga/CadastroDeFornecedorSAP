"""
Script de Instalação - Automação SAP Modular
Execute este script no diretório raiz do projeto
"""

import os
import shutil
from pathlib import Path


def instalar_modulo_sap():
    """Instala o módulo SAP no projeto"""
    
    print("="*70)
    print("INSTALAÇÃO - MÓDULO SAP")
    print("="*70)
    
    # 1. Obter diretório raiz do projeto (onde está main.py)
    projeto_dir = Path(__file__).resolve().parent
    
    print(f"\n📁 Diretório do projeto: {projeto_dir}")
    
    # Verificar se é o diretório correto
    if not (projeto_dir / "main.py").exists():
        print("\n❌ ERRO: Este não é o diretório raiz do projeto!")
        print("   Execute este script no mesmo diretório onde está o main.py")
        return False
    
    # 2. Criar pasta SAP/
    sap_dir = projeto_dir / "SAP"
    
    if sap_dir.exists():
        resposta = input("\n⚠️  A pasta SAP/ já existe. Substituir? (s/n): ")
        if resposta.lower() != 's':
            print("❌ Instalação cancelada pelo usuário")
            return False
        
        print("🗑️  Removendo pasta SAP/ antiga...")
        shutil.rmtree(sap_dir)
    
    print("\n📂 Criando pasta SAP/...")
    sap_dir.mkdir(exist_ok=True)
    
    # 3. Lista de arquivos necessários
    arquivos_necessarios = [
        "__init__.py",
        "ConexaoSAP.py",
        "ManipuladorCampos.py",
        "EntrarTransacao.py",
        "PreencherDados.py",
        "ProcessarPagamentos.py",
        "AutomacaoSAP.py"
    ]
    
    # 4. Verificar se arquivos estão disponíveis
    # (assumindo que estão no mesmo diretório do script de instalação)
    arquivos_fonte = Path(__file__).parent
    
    print("\n📋 Copiando arquivos...")
    for arquivo in arquivos_necessarios:
        fonte = arquivos_fonte / arquivo
        destino = sap_dir / arquivo
        
        if fonte.exists():
            shutil.copy2(fonte, destino)
            print(f"   ✅ {arquivo}")
        else:
            print(f"   ⚠️  {arquivo} - NÃO ENCONTRADO (baixe manualmente)")
    
    # 5. Mover campos_sap.json para SAP/
    campos_sap = projeto_dir / "campos_sap.json"
    campos_sap_destino = sap_dir / "campos_sap.json"
    
    if campos_sap.exists():
        print("\n📄 Movendo campos_sap.json para SAP/...")
        shutil.copy2(campos_sap, campos_sap_destino)
        print("   ✅ campos_sap.json copiado")
    else:
        print("\n⚠️  campos_sap.json não encontrado na raiz")
        print("   Copie manualmente para SAP/campos_sap.json")
    
    # 6. Criar arquivo de teste
    print("\n📝 Criando script de teste...")
    
    teste_content = '''"""
Teste de Instalação do Módulo SAP
"""

def testar_instalacao():
    print("="*70)
    print("TESTE DE INSTALAÇÃO - MÓDULO SAP")
    print("="*70)
    
    # Teste 1: Importação básica
    print("\\n[Teste 1] Importando executar_automacao...")
    try:
        from SAP.AutomacaoSAP import executar_automacao
        print("✅ PASSOU - Importação OK")
    except ImportError as e:
        print(f"❌ FALHOU - {e}")
        return False
    
    # Teste 2: Módulos individuais
    print("\\n[Teste 2] Importando módulos individuais...")
    try:
        from SAP.ConexaoSAP import ConexaoSAP
        from SAP.ManipuladorCampos import ManipuladorCamposSAP
        from SAP.EntrarTransacao import EntrarTransacao
        from SAP.PreencherDados import PreencherDados
        from SAP.ProcessarPagamentos import ProcessarPagamentos
        print("✅ PASSOU - Todos os módulos OK")
    except ImportError as e:
        print(f"❌ FALHOU - {e}")
        return False
    
    # Teste 3: Verificar campos_sap.json
    print("\\n[Teste 3] Verificando campos_sap.json...")
    from pathlib import Path
    
    campos_json = Path("SAP/campos_sap.json")
    if campos_json.exists():
        print("✅ PASSOU - campos_sap.json encontrado")
    else:
        print("❌ FALHOU - campos_sap.json não encontrado")
        print("   Copie campos_sap.json para a pasta SAP/")
        return False
    
    print("\\n" + "="*70)
    print("✅ INSTALAÇÃO VERIFICADA COM SUCESSO!")
    print("="*70)
    print("\\n🚀 O módulo está pronto para uso!")
    return True

if __name__ == "__main__":
    testar_instalacao()
'''
    
    teste_file = projeto_dir / "testar_sap.py"
    with open(teste_file, 'w', encoding='utf-8') as f:
        f.write(teste_content)
    
    print(f"   ✅ {teste_file.name} criado")
    
    # 7. Resumo
    print("\n" + "="*70)
    print("✅ INSTALAÇÃO CONCLUÍDA!")
    print("="*70)
    print("\n📁 Estrutura criada:")
    print(f"   {sap_dir}/")
    for arquivo in arquivos_necessarios:
        if (sap_dir / arquivo).exists():
            print(f"      ✅ {arquivo}")
        else:
            print(f"      ❌ {arquivo} - AUSENTE")
    
    if (sap_dir / "campos_sap.json").exists():
        print(f"      ✅ campos_sap.json")
    else:
        print(f"      ⚠️  campos_sap.json - COPIE MANUALMENTE")
    
    print("\n🧪 Próximo passo:")
    print(f"   python testar_sap.py")
    
    return True


if __name__ == "__main__":
    import sys
    
    sucesso = instalar_modulo_sap()
    
    if sucesso:
        print("\n✅ Execute 'python testar_sap.py' para verificar a instalação")
        sys.exit(0)
    else:
        print("\n❌ Instalação falhou. Verifique os erros acima.")
        sys.exit(1)