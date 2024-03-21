/**
 * Posts a new message to the specified space on behalf of the user.
 * @param {string} spaceName The resource name of the space.
 */
function postMessageWithUserCredentials(spaceName) {
  spaceName = 'h1TQhEAAAAE'
  try {
    const message = {'text': 'Hello world!'};
    Chat.Spaces.Messages.create(message, spaceName);
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed to create message with error %s', err.message);
  }
}