"""Extracteur principal"""

import time
from typing import Optional, List
import pdfplumber

from .models import (
    DocumentType, ExtractionResult, IdentificationData,
    TableType, BilanActifLine, BilanPassifLine, CPCLine
)
from .parser_ammc import AMMCParser
from .parser_dgi import DGIParser
from .utils import detect_document_type, calculate_confidence, merge_multipage_tables
from config.settings import ExtractionConfig


class FiscalPDFExtractor:
    """Extracteur principal pour les documents fiscaux marocains"""
    
    def __init__(self, pdf_path, doc_type: str = None, config: ExtractionConfig = None):
        """
        Initialise l'extracteur.
        
        Args:
            pdf_path: Chemin vers le fichier PDF ou objet fichier
            doc_type: Type de document ("AMMC", "DGI", ou None pour auto-détection)
            config: Configuration d'extraction
        """
        self.pdf_path = pdf_path
        self.doc_type = doc_type
        self.config = config or ExtractionConfig()
        
        self.ammc_parser = AMMCParser(self.config)
        self.dgi_parser = DGIParser(self.config)
        
        self.result = None
    
    def extract_all(self, progress_callback=None) -> ExtractionResult:
        """
        Extrait toutes les données du PDF.
        
        Args:
            progress_callback: Fonction de callback pour la progression
            
        Returns:
            ExtractionResult contenant toutes les données
        """
        start_time = time.time()
        
        with pdfplumber.open(self.pdf_path) as pdf:
            # Détection automatique du type si non spécifié
            if not self.doc_type:
                first_page_text = pdf.pages[0].extract_text()
                detected_type = detect_document_type(first_page_text)
                self.doc_type = DocumentType(detected_type) if detected_type != "UNKNOWN" else DocumentType.AMMC
            
            # Sélection du parser
            if self.doc_type == DocumentType.AMMC:
                parser = self.ammc_parser
            else:
                parser = self.dgi_parser
            
            # Extraction des données
            identification = None
            bilan_actif_pages = []
            bilan_passif_pages = []
            cpc_pages = []
            
            total_pages = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(i + 1, total_pages)
                
                # Identification (première page)
                if i == 0:
                    identification = parser.parse_identification(page)
                
                # Détection et extraction des tableaux
                table_type = parser.detect_table_type(page)
                
                if table_type == TableType.BILAN_ACTIF:
                    bilan_actif_pages.append(parser.parse_bilan_actif(page))
                elif table_type == TableType.BILAN_PASSIF:
                    bilan_passif_pages.append(parser.parse_bilan_passif(page))
                elif table_type == TableType.CPC:
                    cpc_pages.append(parser.parse_cpc(page))
            
            # Fusion des données multi-pages
            bilan_actif = self._merge_lines(bilan_actif_pages)
            bilan_passif = self._merge_lines(bilan_passif_pages)
            cpc = self._merge_lines(cpc_pages)
            
            # Calcul du score de confiance
            confidence = self._calculate_global_confidence(
                bilan_actif, bilan_passif, cpc, identification
            )
            
            # Création du résultat
            self.result = ExtractionResult(
                document_type=self.doc_type,
                identification=identification or IdentificationData(),
                bilan_actif=bilan_actif,
                bilan_passif=bilan_passif,
                cpc=cpc,
                extraction_time=time.time() - start_time,
                pages_processed=total_pages,
                tables_found=len(bilan_actif_pages) + len(bilan_passif_pages) + len(cpc_pages),
                confidence_score=confidence
            )
            
            # Validation post-extraction
            self._validate_result()
            
        return self.result
    
    def _merge_lines(self, pages_data: List[List]) -> List:
        """Fusionne les données de plusieurs pages"""
        if not pages_data:
            return []
        
        merged = []
        for page_data in pages_data:
            merged.extend(page_data)
        
        return merged
    
    def _calculate_global_confidence(self, actif, passif, cpc, identification) -> float:
        """Calcule le score de confiance global"""
        scores = []
        
        # Score basé sur la présence de données
        if len(actif) > 10:
            scores.append(0.9)
        elif len(actif) > 5:
            scores.append(0.7)
        elif len(actif) > 0:
            scores.append(0.5)
        
        if len(passif) > 5:
            scores.append(0.9)
        elif len(passif) > 2:
            scores.append(0.7)
        
        if len(cpc) > 10:
            scores.append(0.9)
        elif len(cpc) > 5:
            scores.append(0.7)
        
        if identification.raison_sociale:
            scores.append(1.0)
        elif identification.identifiant_fiscal:
            scores.append(0.8)
        
        if not scores:
            return 0.0
        
        return sum(scores) / len(scores)
    
    def _validate_result(self):
        """Valide le résultat et ajoute des warnings"""
        if not self.result:
            return
        
        # Vérification des postes obligatoires
        required_actif = self.config.REQUIRED_ACCOUNTS["actif"]
        found_actif = [line.designation for line in self.result.bilan_actif]
        
        for required in required_actif:
            if not any(required.lower() in found.lower() for found in found_actif):
                self.result.warnings.append(f"Poste obligatoire non trouvé dans l'actif: {required}")
        
        # Vérification des totaux
        total_lines_actif = [l for l in self.result.bilan_actif if "TOTAL" in l.designation.upper()]
        if not total_lines_actif:
            self.result.warnings.append("Aucune ligne TOTAL trouvée dans le bilan actif")
        
        # Vérification des résultats
        result_lines = [l for l in self.result.cpc if "RESULTAT NET" in l.designation.upper()]
        if not result_lines:
            self.result.warnings.append("Aucune ligne RESULTAT NET trouvée dans le CPC")
        
        # Vérification de la cohérence des totaux
        if len(self.result.bilan_actif) > 0 and len(self.result.bilan_passif) > 0:
            actif_total = None
            passif_total = None
            
            for line in self.result.bilan_actif:
                if "TOTAL GENERAL" in line.designation.upper():
                    actif_total = line.net_n
                    break
            
            for line in self.result.bilan_passif:
                if "TOTAL GENERAL" in line.designation.upper():
                    passif_total = line.exercice_n
                    break
            
            if actif_total and passif_total:
                difference = abs(actif_total - passif_total)
                if difference > 1000:  # Tolérance de 1000 DH
                    self.result.warnings.append(
                        f"Écart entre total actif ({actif_total:,.2f}) et total passif ({passif_total:,.2f})"
                    )
    
    def get_result(self) -> Optional[ExtractionResult]:
        """Retourne le résultat de l'extraction"""
        return self.result
