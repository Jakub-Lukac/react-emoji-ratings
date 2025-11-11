import * as React from 'react';

import styles from './ReactEmojiReactionRating.module.scss';
import { IReactEmojiReactionRatingProps } from './IReactEmojiReactionRatingProps';
import { IReactEmojiReactionRatingState } from './IReactEmojiReactionRatingState';

import { DisplayMode, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  TextField, Stack, Label, PrimaryButton
} from '@microsoft/office-ui-fabric-react-bundle';

import spService from './services/spService';

import Badge from '@material-ui/core/Badge/Badge';

import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react';
import { IRatingNewItem } from './models/IRatingNewItem';
import * as e from 'express';

interface IUserReaction {
  user: string;
  displayName: string;
  comment: string;
}

export default class ReactEmojiReactionRating extends React.Component<IReactEmojiReactionRatingProps, IReactEmojiReactionRatingState> {
  private webPartRef = React.createRef<HTMLDivElement>();
  private _spService: spService = null;
  private _message = null;
  private _currentContext = null;
  private listName = this.props.listName; // will always be populated through manifest.json
  private centralizedSiteUrl = this.props.centralSiteUrl; // will always be populated through manifest.json
  private existingRating: any;
  constructor(prop: IReactEmojiReactionRatingProps, state: IReactEmojiReactionRatingState) {
    super(prop);
    this.state = {
      selectedRatingValueShowRections: "",
      selectedRatingIndex: null,
      selectedRatingValue: "",
      0: 0,
      1: 0,
      2: 0,
      3: 0,
      4: 0,
      RatingComments: "",
      CustomMessage: "",
      configLoaded: false,
      isDialogHidden: true,
      dialogTitle: "",
      dialogBody: "",
      ratingItems: [], // preserve state of rating items to fetch users who have rated the page
      showUsersModal: false,
      usersReactedList: [], // list of users who reacted with the selected emoji  
    };

    this._currentContext = this.props.context;
    this._spService = new spService(this.props.context);
    if (Environment.type === EnvironmentType.SharePoint) {
      let items = this.getItems();
      console.log("SharePoint items: ", items);
    }
    else if (Environment.type === EnvironmentType.Local) {
      this._message = <div>Whoops! you are using local host...</div>;
    }

    this.selectedRating = this.selectedRating.bind(this);
    this.submitRating = this.submitRating.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this._onConfigure = this._onConfigure.bind(this);
    this.closeDialog = this.closeDialog.bind(this);
  }

  public componentDidMount() {
    if (this.props.enableCount) {
      this.getItems();
    }
    if (this.props.listName && (this.props.emojisCollection.length > 0)) {
      this.setState({ configLoaded: true });
    }
  }

  public componentDidUpdate(previousProps, previousState) {
    if (previousProps.listName !== this.props.listName) {
      if (this.props.listName && (this.props.emojisCollection.length > 0)) {
        this.setState({ configLoaded: true });
        //this.getItems();
      }
    }

    /*if (previousProps.listMessage !== this.props.listMessage) {
      this.ShowDialogMessage("Vytvorenie zoznamu", this.props.listMessage);
    }*/
  }


  private submitRating(event) {

    let ratingCommnets = this.state.RatingComments ? (this.state.RatingComments).trim() : "";
    //let selectedRatingIndex = parseInt(event.target.id);
    //let selectedRatingIndex = event.target.tabIndex;
    // let selectedRatingValue = event.target.title;

    let selectedRatingIndex = this.state.selectedRatingIndex;
    let selectedRatingValue = this.state.selectedRatingValue;
    let ratingField;

    switch (selectedRatingIndex) {
      case 0:
        ratingField = "Rating1";
        break;
      case 1:
        ratingField = "Rating2";
        break;
      case 2:
        ratingField = "Rating3";
        break;
      case 3:
        ratingField = "Rating4";
        break;
      case 4:
        ratingField = "Rating5";
        break;

    }

    if (!(ratingField)) {
      console.log("Something went wrong! Please check with Admin.");
      return false;
    }

    //let pageName = window.location.pathname.substring(window.location.pathname.lastIndexOf("/") + 1);
    /*  let pageName = window.location.href;
     let body: IRatingNewItem = {
       Title: this._currentContext.pageContext.user.displayName,
       Pagename: pageName ? pageName : window.location.href,
       User: this._currentContext?.pageContext.user.loginName,
       Comment: this.props.enableComments ? ratingCommnets : "",
 
     }
     //adding the right rating column and value
     body[ratingColumn] = selectedRatingValue;
     console.log("body object is: ", body); */

    let pageName = window.location.href;
    let pageId = this._currentContext.pageContext.listItem?.uniqueId; // returns string, or return listItem?.id
    console.log("Page context", this._currentContext.pageContext);
    console.log("Page ID", pageId);
    console.log(typeof pageId);
    let body: IRatingNewItem = {
      Title: this._currentContext.pageContext.user.displayName,
      PageID: pageId.toString(),
      Pagename: pageName,
      User: this._currentContext?.pageContext.user.loginName,
      Rating1: "",
      Rating2: "",
      Rating3: "",
      Rating4: "",
      Rating5: "",
    };

    body[ratingField] = selectedRatingValue;
    console.log("body object is: ", body);

    if (!this.existingRating) {
      this._spService.addRatingItem(this.listName, body, this.centralizedSiteUrl)
        .then(value => {
          console.log("Ratting submitted successfully, thank you!", value.data.id);
          this.ShowDialogMessage("Hodnotenie odoslané", "Vaša hodnotenie bolo úspešne odoslané, ďakujeme!");
        })
        .catch(error => {
          console.log("Something went wrong! please contact admin for more information.");
          this.ShowDialogError(error, "Niečo sa pokazilo! pre viac informácií kontaktujte správcu.");
        });
    }
    else {
      this._spService.updateRatingItem(this.listName, body, this.existingRating.Id, this.centralizedSiteUrl)
        .then(value => {
          console.log("Rating updated successfully, thank you!", value.item);
          this.ShowDialogMessage("Hodnotenie aktualizované", "Vaša hodnotenie bolo úspešne aktualizované, ďakujeme!");
        })
        .catch(error => {
          console.log("Something went wrong! please contact admin for more information.");
          this.ShowDialogError(error, "Niečo sa pokazilo! pre viac informácií kontaktujte správcu.");
        });
    }
  }

  private async getItems() {

    console.log("getItems: ", this._currentContext);
    let ratingItems = await this._spService.getListItems(this.listName, this.centralizedSiteUrl);

    this.setState({ ratingItems: ratingItems });

    console.log("ratingItems: ", ratingItems);
    console.log("this.props.emojisCollection: ", this.props.emojisCollection.length);
    console.log("this.props.emojisCollection: ", this.props.emojisCollection);

    // For each emojiColletion imageUrl, retrieve metadata like Title
    for (const emoji of this.props.emojisCollection) {
      await this._spService.getEmojiMetaData(emoji.ImageUrl, this.props.centralSiteUrl).then((meta) => {
        console.log("Image meta data: ", meta);
        emoji.Title = meta.Title;
      }).catch(error => {
        console.log("Error in fetching image metadata: ", error);
      });
    }


    let column1ValIndex, column2ValIndex, column3ValIndex, column4ValIndex, column5ValIndex;

    for (let i = this.props.emojisCollection.length; i > 0; i--) {
      switch (i) {
        case 5:
          column1ValIndex = this.props.emojisCollection.length - i;
          break;
        case 4:
          column2ValIndex = this.props.emojisCollection.length - i;
          break;
        case 3:
          column3ValIndex = this.props.emojisCollection.length - i;
          break;
        case 2:
          column4ValIndex = this.props.emojisCollection.length - i;
          break;
        case 1:
          column5ValIndex = this.props.emojisCollection.length - i;
          break;

      }
    }

    let pageId = this._currentContext.pageContext.listItem?.uniqueId;
    console.log("Current Page ID: ", pageId);
    let pageRatings = await ratingItems.filter((element) => {
      return (element["PageID"] == pageId
      );
    });

    Promise.all([
      this.getRatingCount(pageRatings, 'Rating1', this.props.emojisCollection[column1ValIndex].Title),
      this.getRatingCount(pageRatings, 'Rating2', this.props.emojisCollection[column2ValIndex].Title),
      this.getRatingCount(pageRatings, 'Rating3', this.props.emojisCollection[column3ValIndex].Title),
      this.getRatingCount(pageRatings, 'Rating4', this.props.emojisCollection[column4ValIndex].Title),
      this.getRatingCount(pageRatings, 'Rating5', this.props.emojisCollection[column5ValIndex].Title),
    ])
      .then(results => {
        console.log("countRating1: ", results[0]);
        console.log("countRating2: ", results[1]);
        console.log("countRating3: ", results[2]);
        console.log("countRating4: ", results[3]);
        console.log("countRating5: ", results[4]);

        this.setState(
          {
            0: results[0],
            1: results[1],
            2: results[2],
            3: results[3],
            4: results[4]
          }
        );
      })
      .catch(error => {
        console.log("Error in getting the rating count!", error.message);
      });

    let userLogin = this._currentContext.pageContext.user.loginName;
    let userSelectedRating = await ratingItems.filter((element) => {
      return (element["PageID"] == pageId
        && (element["User"] == userLogin));
    });

    console.log("userSelectedRating: ", userSelectedRating[0]);

    this.existingRating = userSelectedRating[0];
    let currentUserRatingVal = "";
    let currentUserRatingColumn = "";
    //this.props.emojisCollection
    if (userSelectedRating[0]["Rating1"]) {
      currentUserRatingVal = userSelectedRating[0]["Rating1"];
      currentUserRatingColumn = "Rating1";

    }
    else if (userSelectedRating[0]["Rating2"]) {
      currentUserRatingVal = userSelectedRating[0]["Rating2"];
      currentUserRatingColumn = "Rating2";

    }
    else if (userSelectedRating[0]["Rating3"]) {
      currentUserRatingVal = userSelectedRating[0]["Rating3"];
      currentUserRatingColumn = "Rating3";

    }
    else if (userSelectedRating[0]["Rating4"]) {
      currentUserRatingVal = userSelectedRating[0]["Rating4"];
      currentUserRatingColumn = "Rating4";

    }
    else if (userSelectedRating[0]["Rating5"]) {
      currentUserRatingVal = userSelectedRating[0]["Rating5"];
      currentUserRatingColumn = "Rating5";

    }

    let userSelectedRatingIndex;
    await this.props.emojisCollection.filter((element, tabIndex) => {
      if ((element["Title"] == currentUserRatingVal)) {
        userSelectedRatingIndex = tabIndex;
        return tabIndex;
      }
    });
    console.log("userSelectedRatingIndex: ", userSelectedRatingIndex);
    this.setState({
      selectedRatingIndex: userSelectedRatingIndex,
      selectedRatingValue: currentUserRatingVal
    });


  }

  private async getRatingCount(items: any[], colName: string, colValue: string) {
    console.log(`getRatingCount for ${colName} and ${colValue}`, items);

    let ratingCount = await items.filter((element) => {
      return element[colName] == colValue;
    }).length;

    return ratingCount;
  }

  private showReactions(tabIndex: number, ratingValue: string) {
    let selectedRatingIndex = tabIndex;
    let selectedRatingValue = ratingValue;

    console.log(`ShowReactionSelectedRatingIndex: ${selectedRatingIndex}, ShowReactionSelectedRatingValue: ${selectedRatingValue}`);

    this.setState({
      selectedRatingValueShowRections: selectedRatingValue,
      showUsersModal: true, // Always show the modal immediately
      usersReactedList: [] // Reset list initially
    });

    const allRatings = this.state.ratingItems;
    console.log("allRatings: ", allRatings);

    let pageId = this._currentContext.pageContext.listItem?.uniqueId;
    let pageRatings = allRatings.filter((element) => {
      return (element["PageID"] == pageId
      );
    });

    console.log("pageRatings: ", pageRatings);

    if (!pageRatings || pageRatings.length === 0) {
      console.log("No page ratings found.");
    }

    // Ratings columns to check
    const ratingColumns = ['Rating1', 'Rating2', 'Rating3', 'Rating4', 'Rating5'];

    // Find users who reacted with selectedRatingValue in any rating column
    const usersReacted = pageRatings.filter(item =>
      ratingColumns.some(col => item[col] === selectedRatingValue)
    ).map(item => ({
      user: item.Title,
      comment: item.Comments || "",
    }));

    // set the state with the users who reacted
    this.setState({
      usersReactedList: usersReacted,
    });

    console.log(`Users who reacted with emoji "${selectedRatingIndex} - ${selectedRatingValue}":`, usersReacted);
  }


  private selectedRating(event) {
    let selectedRatingIndex = event.target.tabIndex;
    let selectedRatingValue = event.target.title;

    this.setState({
      selectedRatingIndex,
      selectedRatingValue,
    });

    console.log(`SelectedRatingIndex: ${selectedRatingIndex}, SelectedRatingValue: ${selectedRatingValue}`);
  }

  private closeModal = (e?: any) => {
    if (e && e.preventDefault) e.preventDefault();
    this.setState({ showUsersModal: false }, () => {
      if (this.webPartRef.current) {
        this.webPartRef.current.focus({ preventScroll: true });
      }
    });
    console.log("Users modal closed");
  }

  private handleChange(event: any, newValue: string) {
    let partialState = {};
    partialState[event.target.name] = newValue || "";
    this.setState(partialState);
  }

  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }

  private async closeDialog(e: any) {
    if (e && e.preventDefault) e.preventDefault();
    this.setState({ isDialogHidden: true }, () => {
      if (this.webPartRef.current) {
        this.webPartRef.current.focus({ preventScroll: true });
      }
    });
    console.log("Dialog closed");
    console.log("Refreshing items after dialog close");
    setTimeout(() => this.getItems(), 300);
  }

  private ShowDialogMessage(title: string, body: string) {
    this.setState({
      //isRedirectDialogHidden: false,
      isDialogHidden: false,
      dialogTitle: title,
      dialogBody: body,
    });
  }

  private ShowDialogError(
    error: any,
    customErrorMessage: string
  ) {
    if (!customErrorMessage) {
      customErrorMessage = "";
    }
    this.setState({
      isDialogHidden: false,
      dialogTitle: "Chyba: Žiadosť zlyhala.",
      dialogBody: customErrorMessage.concat(
        " Chybová správa: ",
        (error.Message || error.message)
      ),

    });
  }

  public render(): React.ReactElement<IReactEmojiReactionRatingProps> {

    /* if (this.props.listMessage) {
      this.ShowDialogMessage("List Creation", this.props.listMessage);
    } */

    return !(this.state.configLoaded) ? (

      <Placeholder iconName='Edit'
        iconText='Nakonfigurujte webovú časť'
        description='Prosím, nakonfigurujte webovú časť.'
        buttonLabel='Nakonfigurovať'
        hideButton={this.props.displayMode === DisplayMode.Read}
        onConfigure={this._onConfigure} />
    ) :
      (
        <div className={styles.reactEmojiReactionRating}
          style={{ backgroundColor: `${this.props.selectedColor}` }}
          ref={this.webPartRef}
          tabIndex={-1}>
          <div className={styles.container}>
            <div className={styles.row}>
              <Stack>
                <div className={styles.description}>{this.props.ratingText ? this.props.ratingText : ''}</div>
              </Stack>
            </div>

            <div className={styles.row}>
              <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
                {this.props.emojisCollection ? this.props.emojisCollection.map((ratingItem, tabIndex) => (
                  <Stack.Item className={styles.emojiWrapper}>
                    {this.props.enableCount ? (
                      <>
                        <Badge color="secondary" overlap="circular" badgeContent={
                          <span
                            onClick={(e) => {
                              e.stopPropagation(); // Prevent click from bubbling to <img>
                              this.showReactions(tabIndex, ratingItem.Title); // custom event for badge
                            }}
                            style={{ cursor: 'pointer' }}
                          >
                            {this.state[tabIndex]}
                          </span>
                        }>
                          <img src={ratingItem.ImageUrl}
                            className={this.state.selectedRatingIndex == tabIndex ? styles.selectedEmoji : styles.stackImage}
                            style={this.state.selectedRatingIndex == tabIndex ? { backgroundColor: `${this.props.selectedEmojiColor}` } : { backgroundColor: 'rgba(0, 0, 0, 0)' }}
                            title={ratingItem.Title}
                            tabIndex={tabIndex}
                            id={tabIndex.toString()}
                            alt={ratingItem.Title}
                            onClick={this.selectedRating}
                          />
                        </Badge>
                        <Label className={styles.labelClass}>{ratingItem.Title}</Label>
                      </>) : (
                      <>
                        <img src={ratingItem.ImageUrl}
                          className={this.state.selectedRatingIndex == tabIndex ? styles.selectedEmoji : styles.stackImage}
                          title={ratingItem.Title}
                          tabIndex={tabIndex}
                          id={tabIndex.toString()}
                          alt={ratingItem.Title}
                          onClick={this.selectedRating}
                        />
                        <Label className={styles.labelClass}>{ratingItem.Title}</Label>
                      </>
                    )
                    }
                  </Stack.Item>
                )) :
                  <Label>Hodnotiaci zoznam je prázdny...</Label>}
              </Stack>
            </div>

            {this.props.enableComments ? (
              <div className={styles.row}>
                <TextField
                  label={"Pridajte komentár (voliteľné)"}
                  value={this.state.RatingComments}
                  onChange={this.handleChange}
                  name="RatingComments"
                  rows={6}
                  multiline={true}
                  width={80}
                  className={styles.txtArea}
                />
                {this.state.RatingComments}
              </div>

            ) : ""
            }
            {this.state.CustomMessage ? (
              <div className={styles.row}>
                {this.state.CustomMessage}
              </div>
            ) : ""
            }

            {this._message}
            <div className={styles.row}>
              <div className={styles.column10}>
              </div>
              <div className={styles.column2}>
                <PrimaryButton type='submit' text="Odoslať" onClick={this.submitRating} disabled={this.state.selectedRatingValue ? false : true} />
              </div>
            </div>

            <Dialog
              hidden={this.state.isDialogHidden}
              onDismiss={this.closeDialog}
              dialogContentProps={{
                type: DialogType.normal,
                title: this.state.dialogTitle,
                subText: this.state.dialogBody
              }}
              modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
              }}
            >
              <DialogFooter>
                <PrimaryButton type='button' onClick={this.closeDialog} text="OK" />
              </DialogFooter>
            </Dialog>


            <Dialog
              hidden={!this.state.showUsersModal}
              onDismiss={this.closeModal}
              dialogContentProps={{
                type: DialogType.normal,
                title: `Užívatelia, ktorí reagovali s "${this.state.selectedRatingValueShowRections}"`
              }}
              modalProps={{
                isBlocking: false,
                styles: { main: { maxWidth: 400 } },
                elementToFocusOnDismiss: this.webPartRef.current,
              }}
            >
              <div style={{ maxHeight: 250, overflowY: 'auto' }}>
                {this.state.usersReactedList.length > 0 ? (
                  <ul style={{ listStyleType: 'none', padding: 0 }}>
                    {this.state.usersReactedList.map((user, idx) => (
                      <li key={idx} style={{ marginBottom: 3 }}>
                        {user.user}
                      </li>
                    ))}
                  </ul>
                ) : (
                  <div>Žiadni používatelia nájdení pre túto reakciu.</div>
                )}
              </div>
              <DialogFooter>
                <PrimaryButton type='button' onClick={this.closeModal} text="Zavrieť" />
              </DialogFooter>
            </Dialog>
          </div >
        </div >
      );
  }
}