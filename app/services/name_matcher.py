import re
from difflib import SequenceMatcher
from typing import List, Dict, Any, Optional

def normalize_arabic(text: str) -> str:
    """
    Cleans and normalizes Arabic names:
    - Removes punctuation, dots, commas, symbols
    - Removes diacritics / Tashkeel and Tatweel
    - Normalizes Alef forms (أ, إ, آ, ٱ -> ا)
    - Normalizes Ta Marbuta (ة -> ه)
    - Normalizes Ya / Alef Maqsura (ى -> ي)
    - Collapses repeated consecutive letters (e.g. محممد -> محمد, علااا -> علا)
    - Trims and collapses multiple spaces
    """
    if not text:
        return ""
    text = str(text).strip().lower()
    
    # Remove diacritics
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)
    
    # Remove punctuation, symbols, dots, commas
    for ch in [".", ",", "،", "؛", ":", "-", "_", "(", ")", "[", "]", "/", "\\", "\"", "'", "*", "#"]:
        text = text.replace(ch, " ")
        
    # Normalize Alef forms
    text = re.sub(r'[أإآٱ]', 'ا', text)
    # Normalize Ta Marbuta
    text = re.sub(r'[ة]', 'ه', text)
    # Normalize Ya / Alef Maqsura
    text = re.sub(r'[ى]', 'ي', text)
    # Remove Tatweel
    text = re.sub(r'ـ+', '', text)
    
    # Collapse repeated consecutive identical characters (e.g. محممد -> محمد)
    collapsed = []
    for c in text:
        if not collapsed or c != collapsed[-1]:
            collapsed.append(c)
        else:
            # allow at most 1 repetition if needed or collapse completely
            pass
    text = "".join(collapsed)
    
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def calculate_name_similarity(name1: str, name2: str) -> float:
    """
    Calculates similarity score (0.0 to 1.0) between two Arabic names,
    ignoring common titles (دكتور, د, مهندس, فني...).
    """
    n1 = normalize_arabic(name1)
    n2 = normalize_arabic(name2)
    if not n1 or not n2:
        return 0.0
    if n1 == n2:
        return 1.0
        
    tokens1 = set(n1.split())
    tokens2 = set(n2.split())
    
    titles = {"د", "دكتور", "دكتوره", "م", "مهندس", "اخصائي", "اخصائيه", "فني", "فنيه", "ممرض", "ممرضه", "استشاري"}
    t1 = tokens1 - titles
    t2 = tokens2 - titles
    
    # Compare token sets
    if t1 and t2:
        jaccard = len(t1 & t2) / len(t1 | t2)
        # If one is a strict subset of the other (e.g. "عمر خالد" vs "عمر خالد احمد")
        subset_ratio = len(t1 & t2) / min(len(t1), len(t2))
    else:
        jaccard = 0.0
        subset_ratio = 0.0
        
    seq = SequenceMatcher(None, n1, n2).ratio()
    return max(jaccard, subset_ratio * 0.9, seq)

def find_duplicate_names(new_name: str, existing_subscribers: List[Dict[str, Any]], threshold: float = 0.82) -> List[Dict[str, Any]]:
    """
    Scans existing subscribers across all centers/departments for duplicate or similar names.
    Returns matches with similarity >= threshold.
    """
    matches = []
    norm_new = normalize_arabic(new_name)
    if not norm_new:
        return matches
        
    for sub in existing_subscribers:
        existing_name = sub.get("name", "")
        score = calculate_name_similarity(new_name, existing_name)
        if score >= threshold:
            matches.append({
                "existing_id": sub.get("id"),
                "existing_name": existing_name,
                "system_code": sub.get("system_code"),
                "org_name": sub.get("org_name", ""),
                "dept_name": sub.get("dept_name", ""),
                "status": sub.get("status", "ACTIVE"),
                "similarity_score": round(score * 100, 1)
            })
            
    # Sort by highest similarity
    matches.sort(key=lambda x: x["similarity_score"], reverse=True)
    return matches

if __name__ == "__main__":
    test_cases = [
        ("د. أحمد    يوسف...", "أحمدد يوسف"),
        ("د. عمر خالد أحمد", "عمر خالد احمد"),
        ("محممد  علاااا", "محمد علا"),
        ("أحمد علي حسن", "محمد علي حسن"),
    ]
    for n1, n2 in test_cases:
        score = calculate_name_similarity(n1, n2)
        print(f"'{n1}' vs '{n2}' -> Score: {score:.2f}")
