from dataclasses import dataclass

@dataclass
class PhotoMatchResult:
    matchedfilename: str | None
    needsreview: bool = False

def matchphotostospeakers(names, filenames):
    results = []
    remaining = list(filenames)
    for name in names:
        chosen = None
        for fn in remaining:
            if name.lower().replace(" ", "") in fn.lower().replace(" ", ""):
                chosen = fn
                break
        if chosen is None and remaining:
            chosen = remaining[0]
        if chosen in remaining:
            remaining.remove(chosen)
        results.append(PhotoMatchResult(matchedfilename=chosen, needsreview=False))
    return results
