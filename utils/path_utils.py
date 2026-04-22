# -*- coding: utf-8 -*-

def expand_path(template: str, vars: dict) -> str:
    path = template
    for k, v in vars.items():
        path = path.replace(f"{{{k}}}", v)
    return path
