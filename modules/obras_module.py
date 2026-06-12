import re
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

EXTENSOES_VALIDAS = {'.jpg', '.jpeg', '.png'}


class PosteInfo:
    # ─── Preparado para futura categorização de fotos por tipo ─────────────
    # Exemplo de uso futuro:
    #   self.fotos_por_tipo = {
    #       'plaquinha': [],
    #       'barramento': [],
    #       'pe_poste': [],
    #       'vista_geral': [],
    #       'transformador': [],
    #   }

    def __init__(self, pasta: Path):
        self.pasta = pasta
        self.numero = self._extrair_numero()
        self.fotos: list[Path] = []
        self._contar_fotos()

    def _extrair_numero(self) -> int:
        match = re.search(r'poste(\d+)', self.pasta.name, re.IGNORECASE)
        return int(match.group(1)) if match else 0

    def _contar_fotos(self):
        encontradas = set()
        for ext in EXTENSOES_VALIDAS:
            for caminho in self.pasta.glob(f'*{ext}'):
                encontradas.add(caminho)
            for caminho in self.pasta.glob(f'*{ext.upper()}'):
                encontradas.add(caminho)
        self.fotos = sorted(encontradas)

    @property
    def quantidade_fotos(self) -> int:
        return len(self.fotos)

    @property
    def nome_formatado(self) -> str:
        return f'POSTE {self.numero:02d}'

    def __repr__(self):
        return f'PosteInfo({self.pasta.name}, {self.quantidade_fotos} fotos)'


class ObraHandler:
    def __init__(self, pasta_obra: str, projeto_id: str = ''):
        self.pasta_obra = Path(pasta_obra)
        self.projeto_id = projeto_id or self.pasta_obra.name
        self.postes: list[PosteInfo] = []
        self._carregar()

    def _carregar(self):
        if not self.pasta_obra.exists():
            raise FileNotFoundError(f'Pasta não encontrada: {self.pasta_obra}')

        pasta_postes = self.pasta_obra / 'todos_postes'
        if not pasta_postes.exists():
            raise FileNotFoundError(
                f'Pasta "todos_postes" não encontrada em: {self.pasta_obra}'
            )

        logger.info('Pasta todos_postes localizada.')

        for item in sorted(pasta_postes.iterdir()):
            if item.is_dir() and re.match(r'poste\d+', item.name, re.IGNORECASE):
                poste = PosteInfo(item)
                if poste.numero > 0:
                    self.postes.append(poste)
                    logger.info(f'{poste.nome_formatado} encontrado.')

        self.postes.sort(key=lambda p: p.numero)

        if not self.postes:
            raise ValueError(
                f'Nenhuma pasta "poste" encontrada dentro de "todos_postes".'
            )

        logger.info(f'Total de postes: {self.total_postes}')
        logger.info(f'Total de fotos: {self.total_fotos}')

    @property
    def total_postes(self) -> int:
        return len(self.postes)

    @property
    def total_fotos(self) -> int:
        return sum(p.quantidade_fotos for p in self.postes)

    @property
    def nome_projeto(self) -> str:
        return self.pasta_obra.name
