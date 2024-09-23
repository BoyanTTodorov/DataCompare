from app import App
from ttkbootstrap import Window


def main():
    try:
        root = Window(themename='superhero')
        app = App(root)
        root.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")  # Print the error message
    finally:
        print("Application has ended.")  # Code to execute after try/except, regardless of outcome

if __name__ == "__main__":
    main()