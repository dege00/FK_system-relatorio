import re
import logging
from pathlib import Path

from modules.photo_handler import PhotoInfo

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


class VirtualPoste:
    def __init__(self, numero: int, fotos: list[Path]):
        self.numero = numero
        self.fotos = fotos

    @property
    def quantidade_fotos(self) -> int:
        return len(self.fotos)

    @property
    def nome_formatado(self) -> str:
        return f'POSTE {self.numero:02d}'

    def __repr__(self):
        return f'VirtualPoste({self.nome_formatado}, {self.quantidade_fotos} fotos)'


class SequenciaObraHandler:
    def __init__(self, pasta_fotos: str, projeto_id: str = '', fotos_por_poste: int = 4):
        self.pasta_fotos = Path(pasta_fotos)
        self.projeto_id = projeto_id or self.pasta_fotos.name
        self.fotos_por_poste = fotos_por_poste
        self.postes: list[VirtualPoste] = []
        self._carregar()

    def _carregar(self):
        if not self.pasta_fotos.exists():
            raise FileNotFoundError(f'Pasta não encontrada: {self.pasta_fotos}')

        arquivos = set()
        for ext in EXTENSOES_VALIDAS:
            for caminho in self.pasta_fotos.glob(f'*{ext}'):
                arquivos.add(caminho)
            for caminho in self.pasta_fotos.glob(f'*{ext.upper()}'):
                arquivos.add(caminho)

        if not arquivos:
            raise ValueError(f'Nenhuma imagem (JPG/JPEG/PNG) encontrada em: {self.pasta_fotos}')

        fotos_info = [PhotoInfo(arq) for arq in arquivos]
        fotos_info.sort(key=lambda f: f.data_hora)
        fotos_ordenadas = [f.caminho for f in fotos_info]

        logger.info(f'{len(fotos_ordenadas)} fotos carregadas de {self.pasta_fotos}')
        logger.info(f'Fotos por poste: {self.fotos_por_poste}')

        for i in range(0, len(fotos_ordenadas), self.fotos_por_poste):
            grupo = fotos_ordenadas[i:i + self.fotos_por_poste]
            numero_poste = (i // self.fotos_por_poste) + 1
            poste = VirtualPoste(numero_poste, grupo)
            self.postes.append(poste)
            logger.info(f'{poste.nome_formatado}: {poste.quantidade_fotos} fotos')

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
        return self.pasta_fotos.name
