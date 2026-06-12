"""
Custom build hooks — sync tessl-managed skills into package resources.

At build time (or editable install), skills are copied from:
  .tessl/plugins/pyxll/pyxll-agent-skills/skills/<skill>/
into:
  src/pyxll_claude/resources/workspace/skills/<skill>/

If the tessl directory is absent the step is silently skipped so the package
can be built from the git-tracked fallback skills.
"""
import shutil
from pathlib import Path

from setuptools import setup
from setuptools.command.build_py import build_py as _build_py
from setuptools.command.egg_info import egg_info as _egg_info

_TESSL_SKILLS = Path(".tessl/plugins/pyxll/pyxll-agent-skills/skills")
_RESOURCE_SKILLS = Path("src/pyxll_claude/resources/workspace/skills")


def _sync_tessl_skills() -> None:
    if not _TESSL_SKILLS.is_dir():
        return
    for skill_dir in sorted(_TESSL_SKILLS.iterdir()):
        if not skill_dir.is_dir():
            continue
        dest = _RESOURCE_SKILLS / skill_dir.name
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(skill_dir, dest)
        print(f"  synced skill: {skill_dir.name}")


class build_py(_build_py):
    def run(self) -> None:
        _sync_tessl_skills()
        super().run()


class egg_info(_egg_info):
    """egg_info runs during editable installs — keep skills in sync there too."""
    def run(self) -> None:
        _sync_tessl_skills()
        super().run()


setup(cmdclass={"build_py": build_py, "egg_info": egg_info})
