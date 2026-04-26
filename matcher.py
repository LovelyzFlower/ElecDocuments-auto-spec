from sentence_transformers import SentenceTransformer, util
import torch

class SemanticMatcher:
    def __init__(self, model_name='jhgan/ko-sroberta-multitask'):
        # 'jhgan/ko-sroberta-multitask' is a great model for Korean semantic similarity
        # Alternatively, 'all-MiniLM-L6-v2' can be used for general/English
        
        # Check for Apple Silicon MPS or CUDA
        self.device = 'cpu'
        if torch.backends.mps.is_available():
            self.device = 'mps'
        elif torch.cuda.is_available():
            self.device = 'cuda'
            
        print(f"Loading Sentence-Transformer model '{model_name}' on {self.device}...")
        self.model = SentenceTransformer(model_name, device=self.device)
        self.metadata_embeddings = None
        self.metadata_items = []

    def fit_metadata(self, metadata_list):
        """
        Precompute embeddings for the loaded metadata list.
        metadata_list: A list of strings (e.g., column names from the Excel file)
        """
        self.metadata_items = metadata_list
        if not self.metadata_items:
            self.metadata_embeddings = None
            return

        # Encode all metadata entries
        self.metadata_embeddings = self.model.encode(self.metadata_items, convert_to_tensor=True, device=self.device)

    def find_best_matches(self, query_texts, top_k=3):
        """
        Find the best matches for a list of query texts (e.g., OCR results) against the fitted metadata.
        Returns a list of lists of dicts containing the top_k matches and their scores.
        """
        if self.metadata_embeddings is None or len(self.metadata_items) == 0:
            # If no metadata is loaded, return empty matches
            return [[] for _ in query_texts]

        # Encode query texts
        query_embeddings = self.model.encode(query_texts, convert_to_tensor=True, device=self.device)
        
        # Compute cosine similarities
        cos_scores = util.cos_sim(query_embeddings, self.metadata_embeddings)
        
        results = []
        for i in range(len(query_texts)):
            scores = cos_scores[i]
            # Handle cases where top_k > len(metadata)
            k = min(top_k, len(self.metadata_items))
            top_results = torch.topk(scores, k=k)
            
            match_list = []
            for score, idx in zip(top_results[0], top_results[1]):
                match_list.append({
                    "match": self.metadata_items[idx.item()],
                    "score": score.item()
                })
            results.append(match_list)
            
        return results
