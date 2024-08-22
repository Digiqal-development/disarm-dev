import sys

def capitalize_string(input_string):
    capitalized_string = input_string.capitalize()
    return capitalized_string

if __name__ == "__main__":
    # Ensure a string argument is provided
    if len(sys.argv) != 2:
        print("Usage: python capitalize_strincaluseg.py <input_string>")
        sys.exit(1)

    input_string = sys.argv[1]
    print(capitalize_string(input_string))