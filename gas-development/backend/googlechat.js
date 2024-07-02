/**
 * Gửi tin nhắn đến nhân viên.
 * @param {string} spaceName Tên Space.
 */
function postMessageWithUserCredentials(spaceName) {
  spaceName = 'h1TQhEAAAAE'
  try {
    const message = {'text': 'Hello world!'};
    Chat.Spaces.Messages.create(message, spaceName);
  } catch (err) {
    console.log('Failed to create message with error %s', err.message);
  }
}