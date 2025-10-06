// State interface for ReactEmojiReactionRating component
// This interface defines the state structure for the component, including selected ratings, comments, and dialog visibility.
export interface IReactEmojiReactionRatingState {
    selectedRatingValueShowRections: string | null;
    selectedRatingIndex: Number | null;
    selectedRatingValue: string | null;
    0: Number | null;
    1: Number | null;
    2: Number | null;
    3: Number | null;
    4: Number | null;
    RatingComments: string;
    CustomMessage: string;
    configLoaded: boolean;
    isDialogHidden: boolean;
    dialogTitle: string;
    dialogBody: string;
    ratingItems: any[];
    showUsersModal: boolean;
    usersReactedList: { user: string; comment: string }[];
}
