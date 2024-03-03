const TEMPLATE_ID = "REPLACE_WITH_GOOGLE_SLIDE_ID";

const BLACKBAUD_LIST = "REPLACE_WITH_BLACKBAUD_LIST_ID";

const current = SlidesApp.getActivePresentation();
const template = SlidesApp.openById(TEMPLATE_ID).getSlides()[0];

/**
 * Retrieves a list of bithdays from Blackbaud and populates the a Google Slide Show with
 * the retrieved info using a given Google Slide template
 */
function main() {
  console.log("Clearing existing slides...");

  clearSlides_(); // Clear existing slides

  console.log("Slides cleared!");

  console.log("Getting birthday list...");

  const birthdays = filterBirthdays_(getBirthdayList_()); // Filter bithday list

  console.log(`Birthday list retrieved! Found ${birthdays.length} birthday(s)...`);

  console.log("Generating slides...");

  createSlides_(birthdays); // Create slides from template

  console.log("Slides complete!");
}

/**
 * Clears all slides from the current slideshow
 */
function clearSlides_() {
  const slides = current.getSlides();

  slides.forEach(slide => slide.remove());
}

/**
 * Converts a string to title case
 * 
 * @param {string} text The string to convert
 * @return {string} The title case version of the string
 */
function toTitleCase_(text) {
  const parts = text.split(" ");

  parts.map(part => part.charAt(0).toUpperCase() + part.substring(1).toLowerCase());

  return parts.join(" ");
}

/**
 * Filters an array of birthdays to only show the birthdays for today.
 * For Fridays, the birthdays for the weekend are included.
 * 
 * @param {array} birthdays An array of birthdays to filter
 * @return {array} A filtered array
 */
function filterBirthdays_(birthdays) {
  const date = new Date();

  date.setHours(0, 0, 0, 0);

  const endDate = new Date(date);

  if (date.getDay() == 5) {
    endDate.setDate(endDate.getDate() + 2);
  }

  birthdays = birthdays.filter(birthday => {
    const dob = new Date(birthday.dob);

    dob.setFullYear(date.getFullYear());

    return dob >= date && dob <= endDate;
  });

  return birthdays;
}

/**
 * Creates a new slide for each birthday following the specified template
 * 
 * @param {array} birthdays An array of birthdays to generate slides for
 */
function createSlides_(birthdays) {
  for (let i = 0; i < birthdays.length; i++) {
    const birthday = birthdays[i];

    const fullName = toTitleCase_(`${birthday.firstName} ${birthday.lastName}`);
    const dob = birthday.dob;

    let slide;

    try {
      slide = current.appendSlide(template, SlidesApp.SlideLinkingMode.NOT_LINKED);

      const photos = getProfileImage_(birthday.userId);

      if (photos.large_filename_url) {
        const image = slide.getImages().find(image => image.getDescription() == "%profileImage%");

        image.replace("https:" + photos.large_filename_url, true);
      }

      slide.replaceAllText("%name%", `${fullName}, ${dob.getMonth()}/${dob.getDate()}`);
      slide.setSkipped(false);

      console.log(`\tCreated - ${fullName} (${i + 1}/${birthdays.length})`);
    } catch (err) {
      console.error(`\tFailed to create slide for ${fullName}! Skipping...`, err);

      if (slide) {
        try {
          slide.remove();
        } catch (err) {
          console.error(`\tFailed to cleanup slide...`, err);
        }
      }
    }
  }
}

/**
 * Retrieves the list of birthdays from Blackbaud
 * 
 * @return {array} An array containing the details of a birthday
 */
function getBirthdayList_() {
  const blackbaud = getBlackbaudService_();

  let list = getList(blackbaud, BLACKBAUD_LIST);
  
  let items = list.results.rows;

  while (list.count >= 1000) {
    list = getList(blackbaud, BLACKBAUD_LIST, list.page + 1, 1000);

    items = items.concat(list.results.rows);
  }

  items = items.map(item => {
    return {
      userId: item.columns[0].value,
      firstName: item.columns[1].value,
      lastName: item.columns[2].value,
      dob: new Date(item.columns[3].value)
    };
  });

  return items;
}

/**
 * Retrives the profile image of a Blackbaud user
 * 
 * @param {number} userId The ID of the user to get the profile image of
 * @return {array} An array containing the profile image of the user in different sizes
 */
function getProfileImage_(userId) {
  const blackbaud = getBlackbaudService_();

  const user = getUser(blackbaud, userId);

  return user.profile_pictures;
}
