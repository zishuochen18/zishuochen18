#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Unified tool dispatcher for 5 business automation tools.
Allows running each tool from the root directory with centralized logging and status tracking.

Usage:
    python run_tool.py <tool-name> [args...]      Run a specific tool with arguments
    python run_tool.py --status                    Show status of all tools
    python run_tool.py --weekly                    Run all tools marked as 'weekly'
    python run_tool.py --monthly                   Run all tools marked as 'monthly'
    python run_tool.py --list                      List all available tools
"""

import json
import os
import sys
import subprocess
import datetime
from pathlib import Path
from typing import Dict, List, Optional

# Fix encoding for Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class ToolDispatcher:
    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.registry_file = self.root_dir / "tools_registry.json"
        self.logs_dir = self.root_dir / "logs"
        self.history_file = self.root_dir / "run_history.json"

        # Ensure logs directory exists
        self.logs_dir.mkdir(exist_ok=True)

        # Load registry
        self.registry = self._load_registry()
        self.history = self._load_history()

    def _load_registry(self) -> Dict:
        """Load tools registry from JSON."""
        if not self.registry_file.exists():
            print(f"❌ Registry file not found: {self.registry_file}")
            sys.exit(1)

        with open(self.registry_file, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _load_history(self) -> Dict:
        """Load run history from JSON."""
        if not self.history_file.exists():
            return {}

        with open(self.history_file, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _save_history(self):
        """Save run history to JSON."""
        with open(self.history_file, 'w', encoding='utf-8') as f:
            json.dump(self.history, f, ensure_ascii=False, indent=2)

    def list_tools(self):
        """List all available tools."""
        print("\n[Tools] Available Tools:")
        print("=" * 80)
        for name, config in self.registry.items():
            freq = config.get('freq', 'unknown')
            desc = config.get('description', 'N/A')
            print(f"  {name:<20} [{freq:10}] {desc}")
        print("=" * 80 + "\n")

    def show_status(self):
        """Show status of all tools."""
        print("\n[Status] Tool Status:")
        print("=" * 100)
        print(f"{'Tool':<20} {'Freq':<10} {'Last Run':<20} {'Status':<15} {'Duration':<10}")
        print("-" * 100)

        for name, config in self.registry.items():
            freq = config.get('freq', 'unknown')

            if name in self.history:
                hist = self.history[name]
                last_run = hist.get('last_run', 'N/A')
                last_status = hist.get('last_status', '?')
                duration = hist.get('duration', 'N/A')

                status_icon = '[OK]' if last_status == 'success' else '[FAIL]'
                print(f"{name:<20} {freq:<10} {last_run:<20} {status_icon:<15} {duration:<10}")
            else:
                print(f"{name:<20} {freq:<10} {'-':<20} {'[PENDING]':<15} {'-':<10}")

        print("=" * 100 + "\n")

    def run_tool(self, tool_name: str, args: List[str] = None) -> bool:
        """
        Run a specific tool.

        Args:
            tool_name: Name of the tool to run
            args: List of arguments to pass to the tool

        Returns:
            True if successful, False otherwise
        """
        if tool_name not in self.registry:
            print(f"[ERROR] Unknown tool: {tool_name}")
            print(f"\nAvailable tools:")
            for name in self.registry.keys():
                print(f"  - {name}")
            return False

        config = self.registry[tool_name]
        tool_dir = self.root_dir / config['dir']
        entry_script = config['entry']
        script_path = tool_dir / entry_script

        # Verify script exists
        if not script_path.exists():
            print(f"[ERROR] Entry script not found: {script_path}")
            return False

        # Prepare command
        cmd = [sys.executable, entry_script]
        if args:
            cmd.extend(args)

        # Prepare log file
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.logs_dir / f"{tool_name}_{timestamp}.log"

        print(f"\n[RUN] Running: {tool_name}")
        print(f"   Directory: {tool_dir}")
        print(f"   Script: {entry_script}")
        if args:
            print(f"   Args: {' '.join(args)}")
        print(f"   Log: {log_file}")
        print("-" * 80)

        # Run the tool
        start_time = datetime.datetime.now()

        try:
            with open(log_file, 'w', encoding='utf-8') as lf:
                result = subprocess.run(
                    cmd,
                    cwd=tool_dir,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    timeout=3600  # 1 hour timeout
                )

                # Write output to log file
                if result.stdout:
                    lf.write(result.stdout)

                # Also print to console
                if result.stdout:
                    print(result.stdout)

            end_time = datetime.datetime.now()
            duration = str(end_time - start_time)

            # Record result
            is_success = result.returncode == 0
            status = "success" if is_success else "failed"

            self.history[tool_name] = {
                "last_run": start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "last_status": status,
                "duration": duration,
                "exit_code": result.returncode,
                "log_file": str(log_file)
            }
            self._save_history()

            # Print summary
            status_icon = "[OK]" if is_success else "[FAIL]"
            print("-" * 80)
            print(f"{status_icon} {tool_name}: {status.upper()} (Duration: {duration})")

            return is_success

        except subprocess.TimeoutExpired:
            print(f"[ERROR] Tool {tool_name} timed out (exceeded 1 hour)")
            self.history[tool_name] = {
                "last_run": start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "last_status": "timeout",
                "duration": str(datetime.datetime.now() - start_time),
                "exit_code": -1
            }
            self._save_history()
            return False

        except Exception as e:
            print(f"[ERROR] Error running {tool_name}: {e}")
            self.history[tool_name] = {
                "last_run": start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "last_status": "error",
                "duration": str(datetime.datetime.now() - start_time),
                "error": str(e)
            }
            self._save_history()
            return False

    def run_by_frequency(self, freq: str) -> bool:
        """
        Run all tools marked with a specific frequency.

        Args:
            freq: Frequency to run ('weekly', 'monthly', etc.)

        Returns:
            True if all tools succeeded, False otherwise
        """
        tools_to_run = [
            name for name, config in self.registry.items()
            if config.get('freq') == freq
        ]

        if not tools_to_run:
            print(f"[WARN] No tools found with frequency: {freq}")
            return True

        print(f"\n[BATCH] Running {len(tools_to_run)} {freq} tool(s): {', '.join(tools_to_run)}")
        print("=" * 80)

        all_success = True
        for tool_name in tools_to_run:
            success = self.run_tool(tool_name)
            all_success = all_success and success
            print()

        return all_success

    def main(self, args: List[str] = None):
        """Main entry point."""
        if args is None:
            args = sys.argv[1:]

        if not args:
            print(__doc__)
            return

        if args[0] == "--help" or args[0] == "-h":
            print(__doc__)
        elif args[0] == "--list":
            self.list_tools()
        elif args[0] == "--status":
            self.show_status()
        elif args[0] == "--weekly":
            self.run_by_frequency("weekly")
        elif args[0] == "--monthly":
            self.run_by_frequency("monthly")
        else:
            # Run specific tool with arguments
            tool_name = args[0]
            tool_args = args[1:] if len(args) > 1 else None
            self.run_tool(tool_name, tool_args)


if __name__ == "__main__":
    dispatcher = ToolDispatcher()
    dispatcher.main()
