﻿using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;

namespace CGS_DeckGenerator
{
    class DeckItem
    {
        const int RES_MULTIPLIER = 50;

        const int CARD_WID = 7;
        const int CARD_HEI = 10;

        const int HOR_NUM = 10;
        const int VER_NUM = 7;

        public List<CardItem> Deck;

        public DeckItem()
        {
            Deck = new List<CardItem>();
        }

        public void GatherSpreadSheetData(ref SheetsService service, string spreadsheetId, string range, bool containsGrouping)
        {
            Console.WriteLine("Reading range: " + range);

            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            
            //Gather all the data in the first four columns
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Building CardItems. Contains Grouping? " + containsGrouping);
                foreach (IList<Object> column in values)
                {
                    CardItem tempCard = new CardItem(column[0].ToString(), column[1].ToString(), column[2].ToString());
                    if (containsGrouping)
                    {
                        tempCard.Grouping = column[3].ToString();
                    }
                    Deck.Add(tempCard);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        public void DrawDeck(FileInfo pathing)
        {
            try
            {
                const int MAX_CARDS = 70;
                const int border = 5;
                
                int deckIndx;
                Brush borderBrush;

                int width_dimension = CARD_WID * RES_MULTIPLIER;
                int height_dimension = CARD_HEI * RES_MULTIPLIER;

                Console.WriteLine("Attempting to draw deck " + pathing.Name);

                using (Bitmap createdBitmap = new Bitmap(width_dimension * HOR_NUM, height_dimension * VER_NUM))
                {
                    using (Graphics deckImage = Graphics.FromImage(createdBitmap))
                    {
                        for (int v = 0; v < VER_NUM; ++v)
                        {
                            for (int h = 0; h < HOR_NUM; ++h)
                            {
                                //Get deck index
                                deckIndx = (v * HOR_NUM) + h;

                                //Back Fill
                                deckImage.FillRectangle(Brushes.Black, new Rectangle(v * width_dimension, h * height_dimension, width_dimension, height_dimension));

                                //Border
                                borderBrush = GetBorderBrush(Deck[deckIndx].Grouping);
                                deckImage.FillRectangle(Brushes.Green, 
                                    new Rectangle(v * width_dimension, h * height_dimension + border, width_dimension, border)); //top
                                deckImage.FillRectangle(Brushes.White,
                                    new Rectangle(v * width_dimension, h * height_dimension + border, border, height_dimension)); //left
                                deckImage.FillRectangle(Brushes.Yellow,
                                    new Rectangle(v * width_dimension, h * height_dimension + border, width_dimension, height_dimension)); //right
                                deckImage.FillRectangle(Brushes.Blue,
                                    new Rectangle(v * width_dimension, h * height_dimension + border, width_dimension, height_dimension)); //bottom
                            }
                        }
                    }

                    createdBitmap.Save(pathing.FullName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue while trying to draw the deck: " + ex.ToString());
            }
        }

        public Brush GetBorderBrush(string grouping)
        {
            Brush groupingBrush = Brushes.Black;

            switch (grouping)
            {
                case "Common":
                    break;
                default:
                    Console.WriteLine("Found unexpected grouping item: " + grouping);
                    break;
            }

            return groupingBrush;
        }
    }

    public class CardItem
    {
        public string Name;
        public string Effect;
        public string Flavor;
        public string Grouping;

        public CardItem(string n = "", string e = "", string f = "", string g = "")
        {
            Name = n;
            Effect = e;
            Flavor = f;
            Grouping = g;
        }
    }
}
