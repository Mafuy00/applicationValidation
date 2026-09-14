"""
Quick diagnostic: list all Outlook mail stores (mailboxes) and their top-level
folders, so we can identify the shared mailbox / folder to monitor.

Run this with desktop Outlook installed (it will launch/attach to Outlook).
"""
import win32com.client


def walk_folders(folder, depth=0, max_depth=2):
    print("  " * depth + f"- {folder.Name} ({folder.Items.Count} items)")
    if depth >= max_depth:
        return
    for sub in folder.Folders:
        walk_folders(sub, depth + 1, max_depth)


def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    print(f"Found {ns.Folders.Count} store(s) in this Outlook profile:\n")
    for store in ns.Folders:
        print(f"STORE: {store.Name}")
        for folder in store.Folders:
            walk_folders(folder, depth=1, max_depth=2)
        print()


if __name__ == "__main__":
    main()
