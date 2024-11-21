package outlook

import (
	"context"
	"fmt"
	"net/http"
)

// Session manages communication to microsoft's graph api as an authenticated user.
type Session struct {
	client       *Client
	basePath     string
	accessToken  string
	refreshToken string
}

// NewSession returns a new instance of a Session.
func NewSession(client *Client) (*Session, error) {
	token, err := client.tokenSource.Token()
	if err != nil {
		return nil, err
	}

	session := &Session{
		client:       client,
		basePath:     "/me",
		accessToken:  token.AccessToken,
		refreshToken: token.RefreshToken,
	}

	return session, nil
}

func (session *Session) query(ctx context.Context, method, url string, params map[string]interface{}, data interface{}, result interface{}) (*http.Response, error) {
	var queryString string
	if params != nil {
		queryString = createQueryString(params)
	}

	path := fmt.Sprintf("%s%s%s", session.basePath, url, queryString)

	req, err := session.client.NewRequest(ctx, method, path, data)
	if err != nil {
		return nil, err
	}

	if session.accessToken == "" {
		return nil, ErrNoAccessToken
	}

	req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", session.accessToken))

	// May want to detect failures due to invalid or expired tokens, then retry after attempting to refresh the token
	return session.client.Do(ctx, req, result)
}

// Get performs a get request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Get(ctx context.Context, url string, params map[string]interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodGet, url, params, nil, result)
}

// Post performs a post request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Post(ctx context.Context, url string, data interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodPost, url, nil, data, result)
}

// Patch performs a patch request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Patch(ctx context.Context, url string, data interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodPatch, url, nil, data, result)
}

// Delete performs a delete request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Delete(ctx context.Context, url string, params map[string]interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodDelete, url, params, nil, result)
}

// Calendars returns an instance of a CalendarService using this session.
func (session *Session) Calendars() *CalendarService {
	return NewCalendarService(session)
}

// Events returns an instance of a EventService using this session.
func (session *Session) Events() *EventService {
	return NewEventService(session)
}

// Folders returns an instance of a FolderService using this session.
func (session *Session) Folders() *FolderService {
	return NewFolderService(session)
}

// Messages returns an instance of a MessageService using this session.
func (session *Session) Messages() *MessageService {
	return NewMessageService(session)
}
