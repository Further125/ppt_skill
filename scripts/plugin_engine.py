#!/usr/bin/env python3
"""
Plugin engine for ppt-skill.

Supports hook-based plugins in the plugins/ directory.

Hooks:
    pre_build(deck_spec, context) -> modified deck_spec
    post_build(prs, context) -> modified prs
    pre_slide(slide_spec, slide, context) -> modified slide
    post_slide(slide_spec, slide, context) -> modified slide

Usage in build_pptx.py:
    from plugin_engine import load_plugins, run_hooks
    plugins = load_plugins()
    deck_spec = run_hooks(plugins, 'pre_build', deck_spec, context)
"""

import importlib.util
import os
import sys

PLUGIN_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "plugins")


def load_plugins(plugin_dir=None):
    """Load all plugins from the plugin directory."""
    if plugin_dir is None:
        plugin_dir = PLUGIN_DIR

    if not os.path.exists(plugin_dir):
        return []

    plugins = []
    for filename in sorted(os.listdir(plugin_dir)):
        if not filename.endswith(".py") or filename.startswith("_"):
            continue

        path = os.path.join(plugin_dir, filename)
        name = filename[:-3]

        try:
            spec = importlib.util.spec_from_file_location(f"pptskill_plugin_{name}", path)
            module = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = module
            spec.loader.exec_module(module)
            plugins.append({
                "name": name,
                "module": module,
                "path": path
            })
        except Exception as e:
            print(f"Warning: Failed to load plugin '{name}': {e}")

    return plugins


def run_hooks(plugins, hook_name, *args, **kwargs):
    """Run all plugins that implement a given hook."""
    result = args[0] if args else None

    for plugin in plugins:
        hook = getattr(plugin["module"], hook_name, None)
        if hook is None:
            continue

        try:
            hook_result = hook(*args, **kwargs)
            if hook_result is not None:
                result = hook_result
                # Update args for next plugin
                args = (result,) + args[1:]
        except Exception as e:
            print(f"Warning: Plugin '{plugin['name']}' hook '{hook_name}' failed: {e}")

    return result


def list_plugins():
    """List available plugins and their hooks."""
    plugins = load_plugins()
    if not plugins:
        print("No plugins found.")
        return

    print(f"Plugins ({len(plugins)}):")
    for p in plugins:
        hooks = []
        for h in ["pre_build", "post_build", "pre_slide", "post_slide"]:
            if hasattr(p["module"], h):
                hooks.append(h)
        print(f"  {p['name']:20}  hooks: {', '.join(hooks) if hooks else 'none'}")
