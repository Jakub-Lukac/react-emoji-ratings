import { WebPartContext } from "@microsoft/sp-webpart-base";

// IReactEmojiReactionRatingProps interface defines the properties for the ReactEmojiReactionRating component.
// Used in .tsx file (this.props.) to pass data and configuration to the component.
export interface IReactEmojiReactionRatingProps {
  ratingText: string;
  emojisCollection: any[];
  context: WebPartContext;
  enableComments: boolean;
  enableCount: boolean;
  selectedColor: string;
  selectedEmojiColor: string;
  listName: string;
  collectionName: string;
  displayMode: any;
  listMessage: string;
  centralSiteUrl: string;
}
