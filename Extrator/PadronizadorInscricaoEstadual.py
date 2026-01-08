"""
Padronizador de Inscrição Estadual por estado brasileiro.
Aplica formatação específica baseada no estado do fornecedor.
"""
import re
from typing import Tuple, Optional


class PadronizadorInscricaoEstadual:
    """
    Padroniza Inscrição Estadual baseado nas regras de cada estado brasileiro.
    Garante portabilidade e consistência dos dados.
    """
    
    # Regras de formatação por estado (UF -> (tamanho_esperado, descrição))
    REGRAS_ESTADOS = {
        'AC': (13, '13 dígitos (11 + 2 DV)'),
        'AL': (9, '9 dígitos (2 + 6 + 1 DV)'),
        'AP': (9, '9 dígitos (2 + 6 + 1 DV)'),
        'AM': (9, '9 dígitos (2 + 6 + 1 DV)'),
        'BA': ((8, 9), '8 ou 9 dígitos'),  # Aceita ambos
        'CE': (9, '9 dígitos (8 + 1 DV)'),
        'DF': (13, '13 dígitos (11 + 2 DV)'),
        'ES': (9, '9 dígitos (8 + 1 DV)'),
        'GO': (9, '9 dígitos (8 + 1 DV)'),
        'MA': (9, '9 dígitos (8 + 1 DV)'),
        'MT': (11, '11 dígitos (10 + 1 DV)'),
        'MS': (9, '9 dígitos (8 + 1 DV)'),
        'MG': (13, '13 dígitos (3 + 8 + 2 DV)'),
        'PA': (9, '9 dígitos (8 + 1 DV)'),
        'PB': (9, '9 dígitos (8 + 1 DV)'),
        'PR': (10, '10 dígitos (8 + 2 DV)'),
        'PE': (9, '9 dígitos (7 + 2 DV)'),
        'PI': (9, '9 dígitos (8 + 1 DV)'),
        'RJ': (8, '8 dígitos (7 + 1 DV)'),
        'RN': (10, '10 dígitos (9 + 1 DV)'),
        'RS': (10, '10 dígitos'),
        'RO': (13, '13 dígitos (após 01/08/2000)'),  # Antiga: 14 dígitos
        'RR': (9, '9 dígitos (8 + 1 DV)'),
        'SC': (9, '9 dígitos (8 + 1 DV)'),
        'SP': (12, '12 dígitos (com 2 DV)'),
        'SE': (9, '9 dígitos (8 + 1 DV)'),
        'TO': (9, '9 dígitos (8 + 1 DV)')
    }
    
    @staticmethod
    def extrair_digitos(inscricao: str) -> str:
        """
        Extrai apenas dígitos da inscrição estadual.
        
        Args:
            inscricao: Inscrição estadual com ou sem formatação
            
        Returns:
            String contendo apenas dígitos
        """
        if not inscricao:
            return ""
        
        return re.sub(r'\D', '', str(inscricao))
    
    @staticmethod
    def validar_e_ajustar_tamanho(digitos: str, estado: str) -> Tuple[bool, str, Optional[str]]:
        """
        Valida e ajusta a quantidade de dígitos para o estado.
        Adiciona zeros à esquerda se necessário.
        
        Args:
            digitos: String contendo apenas dígitos
            estado: Sigla do estado (UF)
            
        Returns:
            Tupla (válido, digitos_ajustados, mensagem_aviso)
        """
        if estado not in PadronizadorInscricaoEstadual.REGRAS_ESTADOS:
            return False, digitos, f"Estado '{estado}' não encontrado nas regras"
        
        regra = PadronizadorInscricaoEstadual.REGRAS_ESTADOS[estado]
        tamanho_esperado = regra[0]
        descricao = regra[1]
        
        tamanho_atual = len(digitos)
        
        # Bahia aceita 8 ou 9 dígitos
        if isinstance(tamanho_esperado, tuple):
            tamanho_min, tamanho_max = tamanho_esperado
            
            # Se já está no tamanho válido, retorna
            if tamanho_atual in tamanho_esperado:
                return True, digitos, None
            
            # Se tem menos que o mínimo, completa com zeros até o mínimo
            if tamanho_atual < tamanho_min:
                digitos_ajustados = digitos.zfill(tamanho_min)
                aviso = f"IE de {estado} tinha {tamanho_atual} dígitos, completado com zeros à esquerda para {tamanho_min} dígitos"
                return True, digitos_ajustados, aviso
            
            # Se tem mais que o máximo, é erro
            if tamanho_atual > tamanho_max:
                return False, digitos, f"IE de {estado} deve ter {descricao}, mas tem {tamanho_atual} dígitos (excede máximo)"
            
            # Entre min e max, retorna como está
            return True, digitos, None
        
        # Outros estados têm tamanho fixo
        
        # Se já está no tamanho correto, retorna
        if tamanho_atual == tamanho_esperado:
            return True, digitos, None
        
        # Se tem MENOS dígitos, completa com zeros à esquerda
        if tamanho_atual < tamanho_esperado:
            digitos_ajustados = digitos.zfill(tamanho_esperado)
            aviso = f"IE de {estado} tinha {tamanho_atual} dígitos, completado com zeros à esquerda para {tamanho_esperado} dígitos"
            return True, digitos_ajustados, aviso
        
        # Se tem MAIS dígitos, é erro
        return False, digitos, f"IE de {estado} deve ter {descricao}, mas tem {tamanho_atual} dígitos (excede limite)"
    
    @staticmethod
    def formatar_ie_generico(digitos: str, estado: str) -> str:
        """
        Aplica formatação genérica padrão para inscrição estadual.
        
        Regra: Mantém apenas dígitos sem separadores.
        Isso garante compatibilidade com SAP e outros sistemas.
        
        Args:
            digitos: String contendo apenas dígitos
            estado: Sigla do estado (UF)
            
        Returns:
            Inscrição formatada (apenas dígitos)
        """
        # Para máxima portabilidade, mantemos apenas dígitos
        # Formatações específicas (pontos, traços) podem causar problemas
        return digitos
    
    @staticmethod
    def padronizar(inscricao: str, estado: str, permitir_vazio: bool = True) -> Tuple[bool, str, Optional[str]]:
        """
        Padroniza inscrição estadual baseado no estado.
        Adiciona zeros à esquerda se necessário.
        
        Args:
            inscricao: Inscrição estadual original
            estado: Sigla do estado (UF)
            permitir_vazio: Se True, aceita inscrição vazia sem erro
            
        Returns:
            Tupla (sucesso, inscricao_padronizada, mensagem_aviso_ou_erro)
        """
        # Normaliza estado
        estado = str(estado).strip().upper()
        
        # Valida estado
        if estado not in PadronizadorInscricaoEstadual.REGRAS_ESTADOS:
            return False, "", f"Estado '{estado}' inválido ou não suportado"
        
        # Trata inscrição vazia
        inscricao_str = str(inscricao).strip() if inscricao else ""
        
        if not inscricao_str:
            if permitir_vazio:
                return True, "", None
            else:
                return False, "", "Inscrição estadual obrigatória"
        
        # Trata "ISENTO" ou similar
        if inscricao_str.upper() in ['ISENTO', 'ISENTA', 'DISPENSADO', 'DISPENSADA']:
            return True, "ISENTO", None
        
        # Extrai apenas dígitos
        digitos = PadronizadorInscricaoEstadual.extrair_digitos(inscricao_str)
        
        if not digitos:
            return False, "", "Inscrição estadual não contém dígitos válidos"
        
        # Valida e ajusta tamanho (adiciona zeros à esquerda se necessário)
        valido, digitos_ajustados, mensagem = PadronizadorInscricaoEstadual.validar_e_ajustar_tamanho(
            digitos, estado
        )
        
        if not valido:
            return False, digitos, mensagem  # Retorna erro
        
        # Formata (mantém apenas dígitos para portabilidade)
        ie_formatada = PadronizadorInscricaoEstadual.formatar_ie_generico(digitos_ajustados, estado)
        
        return True, ie_formatada, mensagem  # Retorna sucesso + aviso se houver ajuste
    
    @staticmethod
    def obter_info_estado(estado: str) -> Optional[str]:
        """
        Retorna informações sobre o formato esperado para o estado.
        
        Args:
            estado: Sigla do estado (UF)
            
        Returns:
            Descrição do formato ou None se estado inválido
        """
        estado = str(estado).strip().upper()
        
        if estado in PadronizadorInscricaoEstadual.REGRAS_ESTADOS:
            return PadronizadorInscricaoEstadual.REGRAS_ESTADOS[estado][1]
        
        return None


def padronizar_inscricao_estadual(inscricao: str, estado: str) -> Tuple[bool, str, Optional[str]]:
    """
    Função de conveniência para padronização de IE.
    
    Args:
        inscricao: Inscrição estadual original
        estado: Sigla do estado (UF)
        
    Returns:
        Tupla (sucesso, inscricao_padronizada, mensagem_erro)
    """
    return PadronizadorInscricaoEstadual.padronizar(inscricao, estado)


# Teste rápido
if __name__ == "__main__":
    print("="*70)
    print("TESTE DE PADRONIZAÇÃO DE INSCRIÇÃO ESTADUAL")
    print("="*70)
    
    # Casos de teste
    testes = [
        ("0032478230054", "MG", "MG - 13 dígitos"),
        ("123456789", "SP", "SP - 12 dígitos (deve falhar)"),
        ("123456789012", "SP", "SP - 12 dígitos"),
        ("12345678", "RJ", "RJ - 8 dígitos"),
        ("123456789", "CE", "CE - 9 dígitos"),
        ("", "MG", "MG - vazio (deve passar)"),
        ("ISENTO", "SP", "SP - ISENTO"),
        ("12.345.678-9", "RJ", "RJ - 8 dígitos com formatação"),
    ]
    
    for inscricao, estado, descricao in testes:
        print(f"\n{descricao}")
        print(f"  Entrada: '{inscricao}' (Estado: {estado})")
        
        sucesso, resultado, erro = padronizar_inscricao_estadual(inscricao, estado)
        
        if sucesso:
            print(f"  ✅ Resultado: '{resultado}'")
        else:
            print(f"  ❌ Erro: {erro}")
            print(f"  Parcial: '{resultado}'")