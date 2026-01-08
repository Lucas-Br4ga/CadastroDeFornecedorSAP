"""
Validadores e formatadores para campos de entrada.
Fornece máscaras automáticas e validação de dados.
"""
import re
from PySide6.QtGui import QValidator
from PySide6.QtCore import Qt


class CNPJValidator(QValidator):
    """Validador para CNPJ com máscara automática"""
    
    def validate(self, text, pos):
        # Remove caracteres não numéricos
        numbers = re.sub(r'\D', '', text)
        
        # Limita a 14 dígitos
        if len(numbers) > 14:
            return (QValidator.Invalid, text, pos)
        
        return (QValidator.Acceptable, text, pos)
    
    @staticmethod
    def format(text):
        """Formata CNPJ: 00.000.000/0000-00"""
        numbers = re.sub(r'\D', '', text)
        
        if len(numbers) <= 2:
            return numbers
        elif len(numbers) <= 5:
            return f"{numbers[:2]}.{numbers[2:]}"
        elif len(numbers) <= 8:
            return f"{numbers[:2]}.{numbers[2:5]}.{numbers[5:]}"
        elif len(numbers) <= 12:
            return f"{numbers[:2]}.{numbers[2:5]}.{numbers[5:8]}/{numbers[8:]}"
        else:
            return f"{numbers[:2]}.{numbers[2:5]}.{numbers[5:8]}/{numbers[8:12]}-{numbers[12:14]}"


class CEPValidator(QValidator):
    """Validador para CEP com máscara automática"""
    
    def validate(self, text, pos):
        numbers = re.sub(r'\D', '', text)
        
        if len(numbers) > 8:
            return (QValidator.Invalid, text, pos)
        
        return (QValidator.Acceptable, text, pos)
    
    @staticmethod
    def format(text):
        """Formata CEP: 00000-000"""
        numbers = re.sub(r'\D', '', text)
        
        if len(numbers) <= 5:
            return numbers
        else:
            return f"{numbers[:5]}-{numbers[5:8]}"


class TelefoneValidator(QValidator):
    """Validador para telefone/celular com máscara automática"""
    
    def validate(self, text, pos):
        numbers = re.sub(r'\D', '', text)
        
        # Permite até 11 dígitos (celular com 9)
        if len(numbers) > 11:
            return (QValidator.Invalid, text, pos)
        
        return (QValidator.Acceptable, text, pos)
    
    @staticmethod
    def format(text):
        """Formata telefone: (00) 0000-0000 ou (00) 00000-0000"""
        numbers = re.sub(r'\D', '', text)
        
        if len(numbers) <= 2:
            return f"({numbers}" if numbers else ""
        elif len(numbers) <= 6:
            return f"({numbers[:2]}) {numbers[2:]}"
        elif len(numbers) <= 10:
            return f"({numbers[:2]}) {numbers[2:6]}-{numbers[6:]}"
        else:
            # Celular com 9 dígitos
            return f"({numbers[:2]}) {numbers[2:7]}-{numbers[7:11]}"


class NumeroValidator(QValidator):
    """Validador para números inteiros positivos"""
    
    def validate(self, text, pos):
        if text == "" or text.isdigit():
            return (QValidator.Acceptable, text, pos)
        return (QValidator.Invalid, text, pos)


class TextoLimitadoValidator(QValidator):
    """Validador para texto com limite de caracteres"""
    
    def __init__(self, max_length: int):
        super().__init__()
        self.max_length = max_length
    
    def validate(self, text, pos):
        if len(text) <= self.max_length:
            return (QValidator.Acceptable, text, pos)
        return (QValidator.Invalid, text, pos)


class CodigoBancoValidator(QValidator):
    """Validador para código do banco (máximo 3 dígitos)"""
    
    def validate(self, text, pos):
        # Permite vazio ou até 3 dígitos
        if text == "":
            return (QValidator.Acceptable, text, pos)
        
        if text.isdigit() and len(text) <= 3:
            return (QValidator.Acceptable, text, pos)
        
        return (QValidator.Invalid, text, pos)


class AgenciaLimitadaValidator(QValidator):
    """Validador para agência bancária (máximo 4 dígitos)"""
    
    def validate(self, text, pos):
        # Permite números, pontos, hífens, mas limita o tamanho
        if not re.match(r'^[\d\.\-]*$', text):
            return (QValidator.Invalid, text, pos)
        
        # Remove caracteres não numéricos para verificar tamanho
        numeros = re.sub(r'\D', '', text)
        if len(numeros) <= 4:
            return (QValidator.Acceptable, text, pos)
        
        return (QValidator.Invalid, text, pos)


class EmailValidator(QValidator):
    """Validador básico para email"""
    
    EMAIL_REGEX = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
    
    def validate(self, text, pos):
        if text == "":
            return (QValidator.Intermediate, text, pos)
        
        if self.EMAIL_REGEX.match(text):
            return (QValidator.Acceptable, text, pos)
        
        # Permite digitação parcial
        return (QValidator.Intermediate, text, pos)
    
    @staticmethod
    def is_valid(email):
        """Verifica se o email é válido"""
        return bool(EmailValidator.EMAIL_REGEX.match(email))


class ContaBancariaValidator(QValidator):
    """Validador para conta bancária (permite números, hífen e dígito verificador)"""
    
    def validate(self, text, pos):
        # Permite números, hífen e espaços
        if re.match(r'^[\d\-\s]*$', text):
            return (QValidator.Acceptable, text, pos)
        return (QValidator.Invalid, text, pos)


class AgenciaValidator(QValidator):
    """Validador para agência bancária"""
    
    def validate(self, text, pos):
        # Permite números, pontos e hífens
        if re.match(r'^[\d\.\-]*$', text):
            return (QValidator.Acceptable, text, pos)
        return (QValidator.Invalid, text, pos)


def aplicar_validador(campo, tipo_campo):
    """
    Aplica o validador apropriado ao campo baseado no tipo.
    
    Args:
        campo: QLineEdit onde será aplicado o validador
        tipo_campo: string identificando o tipo ('cnpj', 'cep', 'telefone', etc.)
    """
    validadores = {
        'cnpj': CNPJValidator,
        'cep': CEPValidator,
        'telefone': TelefoneValidator,
        'numero': NumeroValidator,
        'email': EmailValidator,
        'conta': ContaBancariaValidator,
        'agencia': AgenciaValidator,
        'agencia_limitada': AgenciaLimitadaValidator,
        'codigo_banco': CodigoBancoValidator,
    }
    
    # Validadores com parâmetros
    if tipo_campo == 'razao_social':
        campo.setValidator(TextoLimitadoValidator(40))
    elif tipo_campo == 'nome_fantasia':
        campo.setValidator(TextoLimitadoValidator(20))
    elif tipo_campo in validadores:
        campo.setValidator(validadores[tipo_campo]())


def aplicar_mascara_automatica(campo, tipo_campo):
    """
    Configura formatação automática enquanto o usuário digita.
    
    Args:
        campo: QLineEdit onde será aplicada a máscara
        tipo_campo: string identificando o tipo ('cnpj', 'cep', 'telefone')
    """
    formatadores = {
        'cnpj': CNPJValidator.format,
        'cep': CEPValidator.format,
        'telefone': TelefoneValidator.format,
    }
    
    if tipo_campo in formatadores:
        formatador = formatadores[tipo_campo]
        
        def formatar_texto():
            cursor_pos = campo.cursorPosition()
            texto_atual = campo.text()
            texto_formatado = formatador(texto_atual)
            
            # Evita loop infinito
            if texto_formatado != texto_atual:
                campo.setText(texto_formatado)
                # Ajusta posição do cursor
                campo.setCursorPosition(min(cursor_pos, len(texto_formatado)))
        
        campo.textChanged.connect(formatar_texto)