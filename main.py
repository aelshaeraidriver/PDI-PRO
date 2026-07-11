import customtkinter as ctk

from gui_v2 import Application


def main():

    ctk.set_appearance_mode(
        "dark"
    )

    ctk.set_default_color_theme(
        "blue"
    )

    root = ctk.CTk()

    Application(root)

    root.mainloop()


if __name__ == "__main__":
    main()