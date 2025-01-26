import * as React from 'react';
import styles from "../../components/Birthdays/Birthdays.module.scss"

import type { IBirthdaysProps } from './IBirthdaysProps';
import "./../workbench.css";
import { IEmailProperties } from "@pnp/sp/sputilities";
import { UncontrolledPopover } from 'reactstrap';
import { Input } from '@material-ui/core';
import Tooltip from "@material-ui/core/Tooltip";
import TextField from '@mui/material/TextField';
import SendIcon from '@mui/icons-material/Send';
import Spotlight from '../Spotlight/Spotlight.cmp';
import Pagination from '@mui/material/Pagination';
import { PaginationItem } from '@mui/material';
import NavigateNextIcon from '@mui/icons-material/NavigateNext';
import NavigateBeforeIcon from '@mui/icons-material/NavigateBefore';
import * as moment from 'moment';
const DefaultProfilePic = require("../../assets/defaultProfilePic.jpg")

export interface IBirthdaysState {
  EvLocation: string;
  EmailText: string;
  IsSent: any;
  BirthdaysList: Array<any>;
  IsLoading: boolean;
  IsModal: boolean;
  People: any;
  openPopoverIndex: number;
  currentPage: number;
  isChangingPage: boolean;
}
export default class Birthdays extends React.Component<IBirthdaysProps, IBirthdaysState> {

  constructor(props: IBirthdaysProps) {
    super(props)
    this.state = {
      BirthdaysList: [],
      EvLocation: '',
      EmailText: '',
      IsSent: null,
      IsLoading: true,
      IsModal: false,
      People: null,
      openPopoverIndex: -1,
      currentPage: 1,
      isChangingPage: false
    }
  }

  componentDidMount(): void {
    this.getBirthdayList();
    document.addEventListener('click', this.handleDocumentClick);
  }
  componentWillUnmount(): void {
    // Remove click listener
    document.removeEventListener('click', this.handleDocumentClick);
  }
  handleDocumentClick = (event: MouseEvent): void => {
    // Convert target to Node for TypeScript compatibility, if necessary
    const target = event.target as Node;

    // Check if click was outside of an open popover
    if (this.state.openPopoverIndex !== -1) {
      // Close popover
      this.setState({ IsModal: false, openPopoverIndex: -1 });
    }
  };
  getBirthdayList = async () => {
    try {
      const date = new Date();
      const rangeStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const rangeEnd = new Date(rangeStart);
      rangeEnd.setDate(rangeEnd.getDate() + parseInt(this.props.BirthdaysRange));

      // get birthdays from sharepoint list
      const items: any[] = await this.props.sp.web.lists.getById(this.props.BirthdayListId)
        .items
        .filter(`IsDisplayed ne 'false'`)
        .orderBy('BirthdayDate')
        .getAll
        ();

      // Filter items based on month and day within the specified range
      const filteredItems = items.filter(item => {
        const birthday = new Date(item.BirthdayDate);
        const birthdayThisYear = new Date(date.getFullYear(), birthday.getMonth(), birthday.getDate());
        return birthdayThisYear >= rangeStart && birthdayThisYear <= rangeEnd;
      });

      // Fetch additional details for each filtered item
      const detailedItems = filteredItems.map(async i => {
        i.isOpen = false;

        try {
          // const userProfile = await this.props.sp.profiles.getPropertiesFor(`i:0#.f|membership|${i.EmployeeEmail}`);
          // const pictureUrlEntry = userProfile.UserProfileProperties.find((prop: any) => prop.Key === "PictureURL");
          let profPic = `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${i?.EmployeeEmail}`
          // i.EmployeePhoto = pictureUrlEntry ? pictureUrlEntry.Value : null;
          i.EmployeePhoto = profPic;
        } catch (error) {
          console.error('Error fetching employee photo', error);
          i.EmployeePhoto = ""; // Or handle the error as needed
        }

        return i;
      });

      let res = await Promise.all(detailedItems);
      res = res.sort((a: any, b: any) => {
        // Convert BirthdayDate to Date objects
        const dateA = new Date(a.BirthdayDate);
        const dateB = new Date(b.BirthdayDate);

        // Compare months first
        if (dateA.getMonth() !== dateB.getMonth()) {
          return dateA.getMonth() - dateB.getMonth();
        }

        // If months are the same, compare dates
        return dateA.getDate() - dateB.getDate();
      })

      this.setState({
        BirthdaysList: res,
        IsLoading: false
      });
    } catch (err) {
      console.error('Error:', err);
      this.setState({ IsLoading: false });
    }
  };


  sendEmail = async (email: string): Promise<boolean> => {
    try {
      if (!this.state.EmailText) return false
      const user = await this.props.sp.web.currentUser();
      const rtlSubject = `\u202Bיש לך ברכה מאת ${user.Title}`;
      const emailProps: IEmailProperties = {
        To: [email],
        From: user.Email,
        // Subject: `You've got a new greeting from ${user.Title}! (${user.Email})`,
        Subject: rtlSubject,
        Body: `
        <div dir="rtl" style="text-align: right;">
          ${this.state.EmailText}
        </div>
      `
      };
      await this.props.sp.utility.sendEmail(emailProps);
      this.setState({ IsSent: true })
      return true;
    } catch (err) {
      console.log('err:', err)
      this.setState({ IsSent: false })
      return false;
    } finally {
      setTimeout(() => {
        this.setState({ IsModal: false as any, EmailText: '', openPopoverIndex: -1 }, () => {
          setTimeout(() => {
            this.setState({ IsSent: null })
          }, 500);
        })
      }, 3000)
    }
  }

  getUserEmail = async (id: any) => {
    let user = await this.props.sp.web.getUserById(id)();
    return user;
  };

  convertToSpDate = (ReleventDate: any): string => {
    // Get day,month and year
    let dd = String(ReleventDate.getDate());
    let mm = String(ReleventDate.getMonth() + 1); //January is 0!
    let yyyy = String(ReleventDate.getFullYear());
    if (parseInt(dd) < 10) {
      dd = "0" + dd;
    }
    if (parseInt(mm) < 10) {
      mm = "0" + mm;
    }
    // Create sp date
    let FormattedReleventDate = yyyy + "-" + mm + "-" + dd + "T00:00:00Z";
    return FormattedReleventDate;
  };

  openModal = (ev: any, itemIndex: number, People: any): void => {
    ev.stopPropagation();
    this.setState({
      IsModal: false,
      openPopoverIndex: -1, // Reset this variable
    }, () => {

      setTimeout(() => {
        this.setState({
          IsModal: true,
          People,
          openPopoverIndex: itemIndex, // Update only this variable
        });
      }, 100);
    });
  }

  closeModal = (ev?: any): void => {
    if (ev) ev.stopPropagation();

    this.setState({
      IsModal: false,
      openPopoverIndex: -1, // Reset this variable
    });
  }

  onPictureConverterUrl = (imageFromSP: any, id: number): string => {
    const imageUrl = this.props.context.pageContext.web.absoluteUrl + "/Lists/Birthdays/Attachments/" +
      id + "/" + JSON.parse(imageFromSP)?.fileName
    return imageUrl
  }

  // handlePageChange = (event: any, value: any) => this.setState({ currentPage: value });
  handleNextClick = () => {
    this.setState({ isChangingPage: true }, () => {
      setTimeout(() => {
        this.setState(prevState => {
          const totalPages = Math.ceil(this.state.BirthdaysList.length / this.props.pageItemsNumber);
          let newState: any = {
            isChangingPage: false,
          };

          if (prevState.currentPage < totalPages) {
            newState.currentPage = prevState.currentPage + 1;
          } else {
            // This ensures `currentPage` is not `undefined`
            newState.currentPage = prevState.currentPage;
          }

          return newState;
        });
      }, 200); // Match this delay with your CSS transition time
    });
  };


  // Similar implementation for `handlePrevClick`

  handlePrevClick = () => {
    this.setState({ isChangingPage: true }, () => {
      setTimeout(() => {
        this.setState(prevState => {
          let newState: any = {
            isChangingPage: false,
          };

          if (prevState.currentPage > 1) {
            newState.currentPage = prevState.currentPage - 1;
          } else {
            // Keep the current page if it's the first page,
            // ensuring `currentPage` is never `undefined`
            newState.currentPage = prevState.currentPage;
          }

          return newState;
        });
      }, 200); // Match this delay with your CSS transition time
    });
  };



  public render(): React.ReactElement<IBirthdaysProps> {

    // Calculate start and end index of items on the current page
    const indexOfLastItem = this.state.currentPage * this.props.pageItemsNumber;
    const indexOfFirstItem = indexOfLastItem - this.props.pageItemsNumber;
    const currentItems = this.state.BirthdaysList.slice(indexOfFirstItem, indexOfLastItem);

    return (
      <section className={` ${styles.Birthdays}`}>
        <div
          style={{ backgroundImage: `url(${this.props.BackgroundImage})` }}
          className={`${styles.BirthdaysBackGroundImage}`}
        >
          <div className={`${styles.BirthdaysLeftContainer}`}>
            <div className={`${styles.BirthdaysContainer}`}>
              <div className={`${styles.BirthdaysTitleCon}`}>
                <span className={`${styles.title}`}>
                  ימי הולדת
                </span>
                {/* <img src={require("../assets/baloons.svg")} alt="" /> */}
              </div>
              <div className='pageing' style={{ display: 'flex', justifyContent: 'space-around', position: 'relative' }}>
                <div className={styles.arrowLeft} onClick={this.handleNextClick} style={{ visibility: this.state.currentPage * this.props.pageItemsNumber < this.state.BirthdaysList.length ? 'visible' : 'hidden' }}
                >
                  <NavigateBeforeIcon fontSize="large" />
                </div>
                <div className={`${styles.BirthdaysPersonContainer} ${this.state.isChangingPage ? styles.changing : ''}`}>

                  {!this.state.IsLoading ? <>
                    {this.state.BirthdaysList.length > 0 ? currentItems.map((People, idx) =>
                      <li id={`popover_${idx}`} key={idx} className={`${styles.item}`} onClick={(ev) => this.openModal(ev, idx, People)}>
                        <div style={{ position: "relative", display: "inline-block" }}>
                          {/* Profile Image */}
                          <img
                            src={People?.EmployeePhoto || DefaultProfilePic}
                            alt=""
                            className={styles.peopleImg}
                            style={{
                              width: "100px", // Adjust width as needed
                              height: "100px", // Adjust height as needed
                              borderRadius: "50%", // Make the profile image circular
                            }}
                          />
                          {/* Balloons Image */}
                          {moment().format('MM-DD') === moment(People?.BirthdayDate).format('MM-DD') && <img

                            src={require("../../assets/balloons.svg")}
                            alt=""
                            style={{
                              position: "absolute",
                              top: "0",
                              left: "0",
                              right: "70px",
                              width: "40px", // Adjust width as needed
                              height: "190px", // Adjust height as needed
                            }}
                          />}
                        </div>

                        <Tooltip title={People?.Title}>
                          <span>{People?.Title}</span>
                        </Tooltip>
                        <span>{moment(People?.BirthdayDate).format("DD/MM")}</span>
                        <div onClick={(e) => e.stopPropagation}>

                          <UncontrolledPopover
                            style={{
                              display: 'flex',
                              justifyContent: 'space-between',
                              alignItems: 'center',
                              flexDirection: 'column',
                              width: '272px',
                              height: '230px',
                              backgroundColor: 'rgb(238 237 237 / 80%)',
                              boxShadow: '4px 4px 4px rgba(0, 0, 0, 0.25)',
                              borderRadius: '25px',
                              backdropFilter: 'blur(5.5px)',
                            }}
                            popperClassName={styles.popoverContainer}
                            placement="top-start"
                            target={`popover_${idx}`}
                            isOpen={this.state.openPopoverIndex === idx}
                          >
                            <div className={styles.modalTriangle}></div>
                            {this.state.IsSent === null && <>
                              <span className={styles.modalHeader}>מזל טוב {People.Title}!</span>

                              <img
                                // src={this.onPictureConverterUrl(People.EmployeePhoto, People.Id)}
                                src={People?.EmployeePhoto || DefaultProfilePic}
                                alt=""
                                className={styles.peopleImgInModal}
                              />
                              <div className={styles.inputContainer}>


                                <TextField
                                  id=""
                                  label=""
                                  type='text'
                                  value={this.state.EmailText}
                                  // onKeyDown={(ev) => ev.key === 'Enter' ? this.sendEmail(People.Email) : null}
                                  onChange={(ev) => this.setState({ EmailText: ev.target.value })}
                                  className={styles.placeholderStyle}
                                  autoFocus={true}
                                  variant='standard'
                                  placeholder={`אחל מזל טוב ל${People.Title}`}
                                  multiline
                                  maxRows={2}
                                />
                                <SendIcon className={styles.sendIcon} htmlColor='#2163a2' onClick={() => this.sendEmail(People?.EmployeeEmail)}></SendIcon>
                              </div>
                            </>}
                            {this.state.IsSent === true && <div className={styles.emailSend}>
                              <img className={styles.emailImg} src={require('../../assets/EmailSentV.png')} alt="" />
                              <span className={styles.emailMsg}>אימייל נשלח ל{People.Title}</span>
                            </div>}

                            {this.state.IsSent === false && <div className={styles.emailSend}>
                              <span className={styles.redX}>X</span>
                              <span className={styles.emailMsg}>שגיאה בעת שליחת המייל ל{People.Title}</span>
                            </div>}

                            {/* {this.state.IsSent && } */}

                          </UncontrolledPopover>
                        </div>
                      </li>
                    ) : <span style={{
                      display: 'flex',
                      alignItems: 'center',
                      fontWeight: 600,
                      paddingTop: '4rem'
                    }}>אין תוצאות</span>}
                  </>
                    : <span style={{
                      display: 'flex',
                      alignItems: 'center',
                      fontWeight: 600,
                      paddingTop: '4rem'
                    }}>בטעינה...</span>}

                </div>

                <div className={styles.arrowRight} >
                  <NavigateNextIcon fontSize="large" onClick={this.handlePrevClick} style={{ visibility: this.state.currentPage > 1 ? 'visible' : 'hidden' }} />
                </div>
              </div>
              <div style={{ position: 'relative', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              </div>
            </div>
          </div>
          <div className={`${styles.BirthdaysRightContainer}`}>
            <Spotlight sp={this.props.sp} SpotlightId={this.props.SpotlightId} context={this.props.context} SpotlightInterval={this.props.SpotlightInterval} />
          </div>
        </div>
      </section>
    );
  }
}
