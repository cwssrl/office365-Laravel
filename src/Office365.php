<?php

namespace Cwssrl\Office365;

use Illuminate\Contracts\Container\Container;
use Microsoft\Graph\Generated\Users\Item\MailFolders\Item\Messages\MessagesRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Authentication\Oauth\AuthorizationCodeContext;

class Office365
{
    /** @var GraphServiceClient */
    private $client;

    // private $graph;

    public function __construct(Container $app)
    {
        $tokenRequestContext = new AuthorizationCodeContext(
            config('office365.tenantId'),
            config('office365.appId'),
            config('office365.secret'),
            'authCode',
            config('office365.redirect_url'),
        );
        $this->client = new GraphServiceClient($tokenRequestContext);

        // $this->client = new GenericProvider([
        //     'clientId'                => $config->get('Office365.appId'),
        //     'clientSecret'            => $config->get('Office365.secret'),
        //     'redirectUri'             => $config->get('Office365.redirect_url'),
        //     'urlAuthorize'            => $config->get('Office365.authority') . $config->get('Office365.authority_endpoint'),
        //     'urlAccessToken'          => $config->get('Office365.authority') . $config->get('Office365.authority_token'),
        //     'urlResourceOwnerDetails' => '',
        //     'scopes'                  => $config->get('Office365.scopes'),
        // ]);
    }

    public function login()
    {
        // return $this->client->getAuthorizationUrl();
    }

    // public function getAccessToken($code)
    // {
    //     $accessToken = $this->client->getAccessToken('authorization_code', [
    //         'code' => $code,
    //     ]);

    //     return [
    //         'token'        => $accessToken->getToken(),
    //         'RefreshToken' => $accessToken->getRefreshToken(),
    //         'expires'      => $accessToken->getExpires(),
    //     ];
    // }

    // public function getUserInfo($user_access_token)
    public function getUserInfo()
    {
        // $this->graph->setAccessToken($user_access_token);

        // $user = $this->graph->createRequest('GET', '/me')
        //     ->execute();

        // return $user->getBody();

        return $this->client->me();
    }

    // public function getEmails($user_access_token, $limit = 10)
    public function getEmails($limit = 10)
    {

        // $this->graph->setAccessToken($user_access_token);

        // $messageQueryParams = [
        //     "\$orderby" => "receivedDateTime DESC",
        //     "\$top"     => $limit,
        // ];

        // return $this->graph->createRequest('GET', '/me/mailfolders/inbox/messages?' . http_build_query($messageQueryParams))
        //     ->setReturnType(Model\Message::class)
        //     ->execute();

        $configuration = new MessagesRequestBuilderGetRequestConfiguration();
        $configuration->queryParameters->orderby = ['receivedDateTime DESC'];
        $configuration->queryParameters->top = $limit;

        return $this->client->me()
            ->mailFolders()
            ->byMailFolderId('inbox')
            ->messages()
            ->get($configuration);
    }
}
