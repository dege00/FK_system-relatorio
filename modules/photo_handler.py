"""
Módulo responsável pelo processamento das fotos.
Gerencia leitura, extração de metadados EXIF e ordenação cronológica.
"""

import os
from pathlib import Path
from datetime import datetime
from PIL import Image
import piexif
import logging

logger = logging.getLogger(__name__)

EXTENSOES_VALIDAS = {'.jpg', '.jpeg', '.png'}


class PhotoInfo:
    """
    Representa uma foto com seus metadados.
    """

    def __init__(self, caminho: Path):
        self.caminho = caminho
        self.nome = caminho.name
        self.data_hora = self._obter_data_hora()

    def _obter_data_hora(self) -> datetime:
        """
        Obtém a data/hora da foto a partir do EXIF.
        Se não houver EXIF, utiliza a data de modificação do arquivo.
        """
        try:
            img = Image.open(self.caminho)
            exif_bytes = img.info.get('exif', b'')
            if exif_bytes:
                try:
                    exif_dict = piexif.load(exif_bytes)
                    exif_data = exif_dict.get('Exif', {})

                    dt_original = exif_data.get(piexif.ExifIFD.DateTimeOriginal)
                    if dt_original:
                        return datetime.strptime(dt_original.decode('utf-8'), '%Y:%m:%d %H:%M:%S')

                    dt_digitized = exif_data.get(piexif.ExifIFD.DateTimeDigitized)
                    if dt_digitized:
                        return datetime.strptime(dt_digitized.decode('utf-8'), '%Y:%m:%d %H:%M:%S')
                except Exception:
                    pass  # EXIF malformado, ignora
            img.close()
        except Exception as e:
            logger.debug(f'Sem EXIF em {self.caminho.name}: {e}')

        # Fallback: data de modificação do arquivo
        try:
            timestamp = os.path.getmtime(self.caminho)
            return datetime.fromtimestamp(timestamp)
        except Exception as e:
            logger.warning(f'Erro ao obter data de modificação de {self.caminho.name}: {e}')
            return datetime.now()

    def __repr__(self):
        return f'PhotoInfo({self.nome}, {self.data_hora})'


class PhotoHandler:
    """
    Gerencia a leitura e ordenação das fotos em uma pasta.
    """

    def __init__(self, pasta_fotos: str):
        self.pasta = Path(pasta_fotos)
        self.fotos: list[PhotoInfo] = []
        self._carregar()

    def _carregar(self):
        """Carrega e ordena todas as fotos da pasta."""
        if not self.pasta.exists():
            raise FileNotFoundError(f'Pasta não encontrada: {self.pasta}')

        arquivos = set()
        for ext in EXTENSOES_VALIDAS:
            for caminho in self.pasta.glob(f'*{ext}'):
                arquivos.add(caminho)
            for caminho in self.pasta.glob(f'*{ext.upper()}'):
                arquivos.add(caminho)

        if not arquivos:
            raise ValueError(f'Nenhuma imagem (JPG/JPEG/PNG) encontrada em: {self.pasta}')

        self.fotos = [PhotoInfo(arq) for arq in arquivos]

        # Ordena da mais antiga para a mais recente
        self.fotos.sort(key=lambda f: f.data_hora)

        logger.info(f'{len(self.fotos)} fotos carregadas de {self.pasta}')

    @property
    def quantidade(self) -> int:
        return len(self.fotos)

    def get_foto(self, indice: int) -> PhotoInfo:
        """Retorna a foto no índice especificado."""
        if 0 <= indice < len(self.fotos):
            return self.fotos[indice]
        raise IndexError(f'Índice {indice} fora do intervalo (0-{len(self.fotos) - 1})')

    def listar_fotos(self) -> list[PhotoInfo]:
        """Retorna a lista completa de fotos ordenadas."""
        return self.fotos.copy()
