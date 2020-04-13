using Google.Apis.Sheets.v4;
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
                const int MAX_CARDS = 69;
                const int border = 20;
                
                int deckIndx;
                Brush borderBrush;

                Font nameFont = new Font("Arial", 16, FontStyle.Bold);
                Font effectFont = new Font("Arial", 14, FontStyle.Regular);
                Font flavorFont = new Font("Arial", 10, FontStyle.Italic);

                int width_dimension = CARD_WID * RES_MULTIPLIER;
                int height_dimension = CARD_HEI * RES_MULTIPLIER;

                if (Deck.Count <= MAX_CARDS)
                {
                    Console.WriteLine("Attempting to draw deck " + pathing.Name);

                    using (Bitmap createdBitmap = new Bitmap(width_dimension * (HOR_NUM + 1), height_dimension * (VER_NUM + 1)))
                    {
                        using (Graphics deckImage = Graphics.FromImage(createdBitmap))
                        {
                            //Vertical
                            for (int v = 0; v < VER_NUM; ++v)
                            {
                                //Horizontal
                                for (int h = 0; h < HOR_NUM; ++h)
                                {
                                    //Get deck index
                                    deckIndx = (v * HOR_NUM) + h;

                                    if (deckIndx < Deck.Count)
                                    {
                                        //Back Fill
                                        deckImage.FillRectangle(Brushes.Black, new Rectangle(h * width_dimension, v * height_dimension, width_dimension, height_dimension));

                                        //Border
                                        borderBrush = GetBorderBrush(Deck[deckIndx].Grouping);

                                        deckImage.FillRectangle(borderBrush,
                                            new Rectangle(h * width_dimension, v * height_dimension, width_dimension, border)); //top
                                        deckImage.FillRectangle(borderBrush,
                                            new Rectangle(h * width_dimension, v * height_dimension, border, height_dimension)); //left
                                        deckImage.FillRectangle(borderBrush,
                                            new Rectangle(((h + 1) * width_dimension) - border, v * height_dimension, border, height_dimension)); //right
                                        deckImage.FillRectangle(borderBrush,
                                            new Rectangle(h * width_dimension, ((v + 1) * height_dimension) - border, width_dimension, border)); //bottom

                                        //Image Fill
                                        deckImage.FillRectangle(Brushes.White,
                                            new Rectangle((h * width_dimension) + (border * 2), (v * height_dimension) + (border * 3),
                                            width_dimension - (border * 4), (height_dimension / 2) - (border * 3)));

                                        //Write Name
                                        RectangleF nameRectangle = new RectangleF(
                                            (h * width_dimension) + (border * 2), (v * height_dimension) + (border * 3 / 2),
                                            width_dimension - (border * 4), (height_dimension / 2) - (border * 3)
                                            );
                                        deckImage.DrawString(Deck[deckIndx].Name, nameFont, Brushes.White, nameRectangle);

                                        //Write Effect
                                        RectangleF effectRectangle = new RectangleF(
                                            (h * width_dimension) + (border * 2), (v * height_dimension) + (height_dimension / 2) + border,
                                            width_dimension - (border * 4), (height_dimension / 2) - (border * 3)
                                            );
                                        deckImage.DrawString(Deck[deckIndx].Effect, effectFont, Brushes.White, effectRectangle);

                                        //Write Flavor
                                        RectangleF eventRectangle = new RectangleF(
                                            (h * width_dimension) + (border * 2), ((v + 1) * height_dimension) - (border * 4),
                                            width_dimension - (border * 4), (height_dimension / 2) - (border * 3)
                                            );
                                        deckImage.DrawString(Deck[deckIndx].Flavor, flavorFont, Brushes.White, eventRectangle);
                                    }
                                }
                            }
                        }

                        createdBitmap.Save(pathing.FullName);
                    }
                }
                else
                {
                    //We'll cross this bridge when we get there
                    Console.WriteLine("Unable to continue, deck is larger than expected max");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue while trying to draw the deck: " + ex.ToString());
            }
        }

        public Brush GetBorderBrush(string grouping)
        {
            Brush groupingBrush = Brushes.Black; //Default

            switch (grouping)
            {
                //Event
                case "":
                    groupingBrush = Brushes.Yellow;
                    break;
                //Resources
                case "Common":
                    groupingBrush = Brushes.White;
                    break;
                case "Uncommon":
                    groupingBrush = Brushes.Green;
                    break;
                case "Rare":
                    groupingBrush = Brushes.Blue;
                    break;
                case "Very Rare":
                    groupingBrush = Brushes.Purple;
                    break;
                //Research
                case "Active":
                    groupingBrush = Brushes.Orange;
                    break;
                case "Passive":
                    groupingBrush = Brushes.Gray;
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
