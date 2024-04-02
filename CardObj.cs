using System.Collections;
using System.Collections.Generic;
using UnityEngine;

// This script handles the creation of Card objects.
public class CardObj : MonoBehaviour
{
    public class Card
    {
        // Each card will have a value and a suit.
        public int value;
        public string suit;

        // Constructor, where the Suit is given as a String.
        public Card(int value, string suit)
        {
            this.value = value;
            this.suit = suit;
        }

        // Constructor, where the Suit is given as an int.
        public Card(int value, int suit)
        {
            this.value = value;
            // Assign suits in the following order:
            // Spades (1) > Hearts (2) > Diamonds (3) > Clubs (4).
            if(suit == 1) {this.suit = "Spades";}
            else if(suit == 2) {this.suit = "Hearts";}
            else if(suit == 3) {this.suit = "Diamonds";}
            else if(suit == 4) {this.suit = "Clubs";}
            // if given a value that isn't an int, or is >4, leave Suit as Null.
        }

        // The Get/Set functions for both the value and suit.
        public int getValue()
        {
            return value;
        }

        public void setValue(int newValue)
        {
            value = newValue;
        }

        public string getSuit()
        {
            return suit;
        }

        public void setSuit (string newSuit)
        {
            suit = newSuit;
        }
    }
}
