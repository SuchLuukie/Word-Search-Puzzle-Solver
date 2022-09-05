# Import files
from extractor import extract_from_excel

class Solver:
    def __init__(self, puzzle, wordlist):
        self.puzzle = puzzle
        self.wordlist = wordlist

        self.adjacencies = [
            (-1, -1), # Top left
            (0, -1),  # Top middle
            (1, -1),  # Top right
            (-1, 0),  # Middle left
            (1, 0),   # Middle right
            (-1, 1),  # Bottom left
            (0, 1),   # Bottom middle
            (1, 1)    # Bottom right
        ]


    # Returns a dictionary of every found word from wordlist and their indices in the 2D puzzle array
    def solve(self):
        found_words = []

        # Divide the words, we're performing a search algorithm on each word
        for word in self.wordlist:
            # Divide it into letters and get the word length
            letters = ([*word])
            indices = self.find_letter_indices(letters[0])

            word_found = False
            # Go through all the indices and try to find the word
            for index in indices:
                # Find the directions of where the next letter is correct
                directions = self.find_next_letter_directions(index, letters[1])
                found_word = [index]

                if word_found:
                    break

                for direction in directions:
                    index = found_word[0]
                    letter_index = 1
                    if word_found:
                        break

                    # Then we look in that direction till we get the wrong letter or till we get to the word length
                    while not word_found:
                        next_letter = letters[letter_index]
                        next_cell = (index[0] + direction[0], index[1] + direction[1])

                        # Only check for the next letter if the index is not negative and is in the board
                        if -1 < next_cell[0] < len(self.puzzle) and -1 < next_cell[1] < len(self.puzzle):
                            # If the next letter is correct we update the variables and look from the current position for the next letter
                            if self.puzzle[next_cell[1]][next_cell[0]] == next_letter:
                                found_word.append(next_cell)
                                index = next_cell
                                letter_index += 1

                                # If the letters we've found is the same length as word, then we've found the word
                                if len(found_word) == len(letters):
                                    found_words.append({word: found_word})
                                    word_found = True
                            
                            # If the next letter is incorrect we break from this branch and try the next
                            else:
                                found_word = [found_word[0]]
                                break

                        else:
                            found_word = [found_word[0]]
                            break

        return found_words
                                

    # Find the directions of where the next letter is positioned
    def find_next_letter_directions(self, index, next_letter):
        directions = []

        # Loop through all the adjencies to find the next letter of the word
        for adj in self.adjacencies:
            next_cell = (index[0] + adj[0], index[1] + adj[1])
            # Only check for the next letter if the index is not negative and is in the board
            if -1 < next_cell[0] < len(self.puzzle) and -1 < next_cell[1] < len(self.puzzle):
                # Check if it's the correct next letter
                if self.puzzle[next_cell[1]][next_cell[0]] == next_letter:
                    directions.append(adj)

        return directions


    # Find all the indices of a letter in the puzzle
    def find_letter_indices(self, letter):
        indices = []
        for y, row in enumerate(self.puzzle):
            for x, cell in enumerate(row):
                if cell == letter:
                    indices.append((x, y))
        
        return indices
                

if __name__ == "__main__":
    puzzle, wordlist = extract_from_excel("puzzles/word_search_puzzles.xlsx", "Woordzoeker 1")

    solver = Solver(puzzle, wordlist)
    solved = solver.solve()

    if len(wordlist) == len(solved):
        print("[!] Puzzle succesfuly solved")
        print("[!] Resuts:\n")
        [print(line) for line in solved]

    else:
        print("[!] Puzzle unsuccesfuly solved")