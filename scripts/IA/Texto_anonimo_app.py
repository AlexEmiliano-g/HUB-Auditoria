# Arquivo: anonimizador_manual.py

import spacy
import re
from collections import defaultdict

class Anonymizer:
    """
    Uma classe para anonimizar texto identificando dados sensíveis (Pessoas, Organizações, CPF, CNPJ)
    e depois reverter essa anonimização.

    O estado (o mapa de dados) é mantido dentro da instância, permitindo
    anonimizar um texto e depois desanonimizar outro (como a resposta de uma IA)
    usando o mesmo mapa.
    """
    def __init__(self):
        print("Inicializando o Anonymizer...")
        try:
            # Carrega o modelo grande em português do spaCy
            self.nlp = spacy.load("pt_core_news_lg")
            print("Modelo de linguagem carregado com sucesso.")
        except OSError:
            print("-" * 50)
            print("ERRO: Modelo 'pt_core_news_lg' não encontrado.")
            print("Por favor, execute no seu terminal: python -m spacy download pt_core_news_lg")
            print("-" * 50)
            exit()
            
        # Dicionário para guardar: {'[PLACEHOLDER]': 'Valor Original'}
        self.mapping = {}
        # Dicionário reverso para evitar criar placeholders duplicados: {'Valor Original': '[PLACEHOLDER]'}
        self.reverse_mapping = {}
        # Contador para gerar placeholders únicos (ex: [PESSOA_1], [PESSOA_2])
        self.counters = defaultdict(int)

        # Padrões de RegEx para dados estruturados
        self.regex_patterns = {
            'CPF': r'\d{3}\.\d{3}\.\d{3}-\d{2}',
            'CNPJ': r'\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}',
            # Adicione outros padrões aqui se precisar (ex: telefone, CEP)
        }

    def _get_placeholder(self, entity_type: str) -> str:
        """Gera um novo placeholder e incrementa o contador para aquele tipo."""
        self.counters[entity_type] += 1
        return f"[{entity_type.upper()}_{self.counters[entity_type]}]"

    def anonymize(self, text: str) -> str:
        """Encontra dados sensíveis no texto e os substitui por placeholders."""
        
        # 1. Análise com spaCy para PER (Pessoa) e ORG (Organização)
        doc = self.nlp(text)
        entities_found = []
        for ent in doc.ents:
            if ent.label_ in ["PER", "ORG"]:
                entities_found.append(ent.text)
        
        # 2. Análise com RegEx
        for entity_type, pattern in self.regex_patterns.items():
            matches = re.findall(pattern, text)
            for match in matches:
                entities_found.append(match)

        # Ordena as entidades encontradas pelo comprimento, da maior para a menor.
        # Isso evita problemas como substituir "João" antes de "João da Silva".
        entities_found.sort(key=len, reverse=True)

        # Cria os placeholders e substitui no texto
        for entity in set(entities_found): # Usamos set() para processar cada entidade única apenas uma vez
            if entity not in self.reverse_mapping:
                # Determina o tipo de entidade para o placeholder
                entity_type = "ENTIDADE" # Padrão
                doc_ent = self.nlp(entity)
                if doc_ent.ents:
                    entity_type = doc_ent.ents[0].label_
                else: # Se o spaCy não pegar, verificamos o regex
                    for type_name, pattern in self.regex_patterns.items():
                        if re.fullmatch(pattern, entity):
                            entity_type = type_name
                            break

                placeholder = self._get_placeholder(entity_type)
                self.mapping[placeholder] = entity
                self.reverse_mapping[entity] = placeholder
            
            # Substitui todas as ocorrências da entidade pelo seu placeholder
            text = text.replace(entity, self.reverse_mapping[entity])
            
        return text

    def deanonymize(self, text: str) -> str:
        """Substitui os placeholders no texto de volta para os valores originais."""
        if not self.mapping:
            print("Aviso: Nenhum mapeamento de anonimização foi criado ainda.")
            return text

        # Ordena os placeholders pelo comprimento para evitar substituições parciais
        sorted_placeholders = sorted(self.mapping.keys(), key=len, reverse=True)

        for placeholder in sorted_placeholders:
            original_value = self.mapping[placeholder]
            text = text.replace(placeholder, original_value)
        return text

def get_multiline_input(prompt: str) -> str:
    """Função auxiliar para capturar múltiplas linhas de texto do usuário."""
    print(prompt)
    lines = []
    while True:
        try:
            line = input()
            if not line:
                break
            lines.append(line)
        except EOFError: # Permite finalizar com Ctrl+D
            break
    return "\n".join(lines)


if __name__ == "__main__":
    print("=" * 50)
    print("### Ferramenta de Anonimização Manual para Auditoria ###")
    print("=" * 50)

    # 1. Cria uma instância da nossa classe
    anonymizer = Anonymizer()

    # 2. Pede o texto original ao usuário
    original_text = get_multiline_input(
        "\nPASSO 1: Cole o texto original abaixo. Pressione Enter em uma linha vazia para continuar.\n"
    )

    # 3. Anonimiza o texto
    anonymized_text = anonymizer.anonymize(original_text)

    # 4. Mostra o texto anonimizado e dá as instruções
    print("\n" + "=" * 50)
    print("TEXTO ANONIMIZADO (PRONTO PARA O GPT):")
    print("-" * 50)
    print(anonymized_text)
    print("=" * 50)
    print("\n>>> AÇÃO: Agora, siga estes passos:")
    print("    1. Copie o texto anonimizado acima.")
    print("    2. Cole no ChatGPT e faça sua pergunta sobre ele.")
    print("    3. Copie a RESPOSTA COMPLETA que o ChatGPT fornecer.")

    # 5. Pede a resposta do GPT
    gpt_response = get_multiline_input(
        "\nPASSO 2: Cole a RESPOSTA ANONIMIZADA do ChatGPT aqui. Pressione Enter em uma linha vazia para finalizar.\n"
    )

    # 6. Desanonimiza a resposta
    final_answer = anonymizer.deanonymize(gpt_response)

    # 7. Mostra o resultado final
    print("\n" + "=" * 50)
    print("RESPOSTA FINAL (DESCRIPTOFRAFADA):")
    print("-" * 50)
    print(final_answer)
    print("=" * 50)